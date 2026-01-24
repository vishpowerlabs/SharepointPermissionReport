import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup, ISharingInfo, ISiteUsage, IRoleDefinitionDetail } from '../models/IPermissionData';

export class PermissionServiceNew implements IPermissionService {
    private readonly _spHttpClient: SPHttpClient;
    private readonly _webUrl: string;

    constructor(spHttpClient: SPHttpClient, webUrl: string) {
        console.log("PermissionServiceNew INITIALIZED - Fresh File!");
        this._spHttpClient = spHttpClient;
        this._webUrl = webUrl;
    }

    public async getSiteRoleAssignments(): Promise<IRoleAssignment[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json && json.value) {
                return json.value.map((item: any) => ({
                    PrincipalId: item.PrincipalId,
                    Member: item.Member,
                    RoleDefinitionBindings: item.RoleDefinitionBindings
                }));
            }
            return [];
        } catch (error) {
            console.error("Error fetching site permissions", error);
            return [];
        }
    }

    public async getLists(excludedLists: string[] = []): Promise<IListInfo[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/lists?$select=Id,Title,ItemCount,Hidden,BaseType,RootFolder/ServerRelativeUrl,RootFolder/StorageMetrics,EntityTypeName,HasUniqueRoleAssignments,RoleAssignments/Member/LoginName,RoleAssignments/Member/PrincipalType,RoleAssignments/RoleDefinitionBindings/Name&$expand=RootFolder/StorageMetrics,RoleAssignments,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`;
            const response = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json && json.value) {
                const lists: IListInfo[] = json.value.map((list: any) => {
                    return {
                        Id: list.Id,
                        Title: list.Title,
                        ItemCount: list.ItemCount,
                        Hidden: list.Hidden,
                        ItemType: list.BaseType === 1 ? "Library" : "List",
                        ServerRelativeUrl: list.RootFolder?.ServerRelativeUrl,
                        HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
                        EntityTypeName: list.EntityTypeName,
                        TotalSize: list.RootFolder?.StorageMetrics?.TotalSize || 0
                    };
                });

                // Filter excludes and hidden system lists (Case Insensitive)
                const normalizedExcludes = excludedLists.map(e => e.toLowerCase());
                return lists.filter(l =>
                    !normalizedExcludes.includes(l.Title.toLowerCase()) &&
                    !l.Hidden &&
                    l.Title !== 'Style Library' && // Common system libs that aren't "Hidden"
                    l.Title !== 'Form Templates' &&
                    l.Title !== 'Site Assets'
                );
            }
            return [];
        } catch (error) {
            console.error("Error fetching lists", error);
            return [];
        }
    }

    public async getListRoleAssignments(listId: string, listTitle: string): Promise<IRoleAssignment[]> {
        try {
            // Use getbyid which is often safer for diverse GUID formats
            const endpoint = `${this._webUrl}/_api/web/lists/getbyid('${listId}')/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json && json.value) {
                return json.value.map((item: any) => ({
                    PrincipalId: item.PrincipalId,
                    Member: item.Member,
                    RoleDefinitionBindings: item.RoleDefinitionBindings
                }));
            }
            return [];
        } catch (error) {
            console.error(`Error fetching permissions for list ${listTitle}`, error);
            return [];
        }
    }

    public async getSiteStats(): Promise<ISiteStats> {
        try {
            const users = await this.getSiteUsers(); // Helper
            const groups = await this.getSiteGroups();
            return {
                totalUsers: users.length,
                totalGroups: groups.length,
                uniquePermissionsCount: 0
            };
        } catch (e) { return { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 }; }
    }

    private async getSiteUsers(): Promise<IUser[]> {
        try {
            const response = await this._spHttpClient.get(`${this._webUrl}/_api/web/siteusers`, SPHttpClient.configurations.v1);
            const json = await response.json();
            return json.value || [];
        } catch { return []; }
    }

    // THE CRITICAL METHOD
    public async getSiteGroups(): Promise<IGroup[]> {
        try {
            // 1. Fetch SharePoint Groups with Users expansion to get count
            const groupsEndpoint = `${this._webUrl}/_api/web/sitegroups?$expand=Users`;
            const groupsReq = this._spHttpClient.get(groupsEndpoint, SPHttpClient.configurations.v1);

            // 2. Fetch Security Groups (PrincipalType 4)
            const secGroupsEndpoint = `${this._webUrl}/_api/web/siteusers?$filter=PrincipalType eq 4`;
            const secGroupsReq = this._spHttpClient.get(secGroupsEndpoint, SPHttpClient.configurations.v1);

            const [groupsResp, secGroupsResp] = await Promise.all([groupsReq, secGroupsReq]);
            const groupsJson = await groupsResp.json();
            const secGroupsJson = await secGroupsResp.json();

            let allGroups: any[] = [];

            // Process SharePoint Groups
            if (groupsJson?.value) {
                const spGroups = groupsJson.value.map((g: any) => {
                    const users = g.Users ? (Array.isArray(g.Users) ? g.Users : (g.Users.value || g.Users.results || [])) : [];
                    return {
                        ...g,
                        PrincipalType: 8,
                        UserCount: users.length // FORCE THIS
                    };
                });
                allGroups = allGroups.concat(spGroups);
            }

            // Process Security Groups
            if (secGroupsJson?.value) {
                const secGroups = secGroupsJson.value.map((g: any) => ({
                    Id: g.Id,
                    Title: g.Title,
                    LoginName: g.LoginName,
                    Description: g.Email || "",
                    OwnerTitle: "Active Directory",
                    IsHiddenInUI: g.IsHiddenInUI,
                    PrincipalType: 4,
                    UserCount: 1 // Assume 1
                }));
                allGroups = allGroups.concat(secGroups);
            }

            return allGroups;

        } catch (error) {
            console.error("Error loading groups in New Service", error);
            return [];
        }
    }

    public async getGroupMembers(groupId: number): Promise<IUser[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/sitegroups/getbyid(${groupId})/users`;
            const response = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();
            return json.value || [];
        } catch { return []; }
    }

    public async searchUsers(query: string): Promise<IUser[]> { return []; }

    public async getSiteAdmins(): Promise<IUser[]> {
        try {
            const response = await this._spHttpClient.get(`${this._webUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`, SPHttpClient.configurations.v1);
            const json = await response.json();
            return json.value || [];
        } catch { return []; }
    }

    public async getSiteOwners(): Promise<IUser[]> {
        try {
            // Fetch users in the associated owners group
            const response = await this._spHttpClient.get(`${this._webUrl}/_api/web/associatedownergroup/users`, SPHttpClient.configurations.v1);
            const json = await response.json();
            return json.value || [];
        } catch { return []; }
    }

    public async getCurrentUser(): Promise<IUser> {
        try {
            const response = await this._spHttpClient.get(`${this._webUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1);
            return await response.json();
        } catch {
            return { Id: 0, Title: "Unknown", IsHiddenInUI: false, PrincipalType: 1 };
        }
    }
    public async removeSitePermission(principalId: number): Promise<boolean> { return false; }
    public async removeListPermission(listId: string, principalId: number): Promise<boolean> { return false; }
    public async removeItemPermission(listId: string, itemId: number, principalId: number): Promise<boolean> { return false; }
    public async removeUserFromGroup(groupId: number, userId: number): Promise<boolean> { return false; }
    public async getUniquePermissionItems(listId: string): Promise<IItemPermission[]> {
        try {
            // Remove OData filter to debug if it's hiding valid items.
            // Fetch top 500 and filter in memory.
            const endpoint = `${this._webUrl}/_api/web/lists/getbyid('${listId}')/items?$select=Id,Title,FileLeafRef,FileRef,HasUniqueRoleAssignments,RoleAssignments/Member/Title,RoleAssignments/Member/LoginName,RoleAssignments/Member/PrincipalType,RoleAssignments/RoleDefinitionBindings/Name&$expand=RoleAssignments,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings&$top=500`;

            const response = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            if (response.ok) {
                const json = await response.json();
                if (json && json.value) {
                    // Filter in memory for unique role assignments
                    const uniqueItems = json.value.filter((i: any) => i.HasUniqueRoleAssignments === true);

                    if (uniqueItems.length === 0 && json.value.length > 0) {
                        console.log("DEBUG: Found items but none marked unique in API response.", json.value[0]);
                    }

                    return uniqueItems.map((item: any) => ({
                        Id: item.Id,
                        Title: item.FileLeafRef || item.Title || "Unknown Item",
                        Url: item.FileRef,
                        RoleAssignments: item.RoleAssignments ? item.RoleAssignments.map((ra: any) => ({
                            PrincipalId: ra.PrincipalId,
                            Member: ra.Member,
                            RoleDefinitionBindings: ra.RoleDefinitionBindings
                        })) : []
                    }));
                }
            }
            return [];
        } catch (e) {
            console.error("Error scanning list for unique items", e);
            return [];
        }
    }
    public async getUserGroups(loginName: string): Promise<IGroup[]> { return []; }
    public async getExternalUsers(): Promise<IUser[]> { return []; }
    public async getSharingLinks(): Promise<ISharingInfo[]> { return []; }
    public async getOrphanedUsers(): Promise<IUser[]> { return []; }
    public async getSiteDetails(): Promise<any> { return { Title: "Site", Url: this._webUrl, HasUniqueRoleAssignments: false }; }
    public async getSiteUsage(): Promise<ISiteUsage> {
        try {
            // Strategy 1: /_api/site/usage
            try {
                const response = await this._spHttpClient.get(`${this._webUrl}/_api/site/usage`, SPHttpClient.configurations.v1);
                if (response.ok) {
                    const json = await response.json();
                    console.log("DEBUG: Site Usage API Response", json);
                    const usage = json.Usage || json.usage;
                    if (usage) {
                        let quota = usage.StorageQuota || usage.storageQuota || 0;
                        const used = usage.Storage || usage.storage || 0;
                        const pct = usage.StoragePercentageUsed || usage.storagePercentageUsed || 0;

                        // Calculate Quota if missing but we have Used and Pct
                        if (quota === 0 && used > 0 && pct > 0) {
                            quota = used / pct;
                            console.log("DEBUG: Calculated Quota from Pct", quota);
                        }

                        return {
                            storageUsed: used,
                            storageQuota: quota,
                            usagePercentage: pct,
                            lastItemModifiedDate: new Date().toISOString()
                        };
                    }
                }
            } catch (e) {
                console.warn("Strategy 1 (site/usage) failed", e);
            }

            // Strategy 2: /_api/site?$select=Usage
            try {
                const response = await this._spHttpClient.get(`${this._webUrl}/_api/site?$select=Usage`, SPHttpClient.configurations.v1);
                if (response.ok) {
                    const json = await response.json();
                    console.log("DEBUG: Site Usage Select Response", json);
                    const usage = json.Usage || json.usage;
                    if (usage) {
                        let quota = usage.StorageQuota || usage.storageQuota || 0;
                        const used = usage.Storage || usage.storage || 0;
                        const pct = usage.StoragePercentageUsed || usage.storagePercentageUsed || 0;

                        // Calculate Quota if missing but we have Used and Pct
                        if (quota === 0 && used > 0 && pct > 0) {
                            quota = used / pct;
                        }

                        return {
                            storageUsed: used,
                            storageQuota: quota,
                            usagePercentage: pct,
                            lastItemModifiedDate: new Date().toISOString()
                        };
                    }
                }
            } catch (e) {
                console.warn("Strategy 2 (site select usage) failed", e);
            }

            return { storageUsed: 0, storageQuota: 0, usagePercentage: 0, lastItemModifiedDate: "" };

        } catch (e) {
            console.error("Error getting site usage", e);
            return { storageUsed: 0, storageQuota: 0, usagePercentage: 0, lastItemModifiedDate: "" };
        }
    }
    public async getRoleDefinitions(): Promise<IRoleDefinitionDetail[]> { return []; }
    public async getUserEffectivePermissions(loginName: string): Promise<string[]> { return []; }
}
