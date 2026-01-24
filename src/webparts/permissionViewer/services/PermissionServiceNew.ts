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
                return json.value.map((item: any) => this._mapRoleAssignment(item));
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
                const systemExcludeSet = new Set(['style library', 'form templates', 'site assets']);
                const userExcludeSet = new Set(excludedLists.map(e => e.toLowerCase()));

                return lists.filter(l => {
                    const titleLower = l.Title.toLowerCase();
                    return !l.Hidden && !systemExcludeSet.has(titleLower) && !userExcludeSet.has(titleLower);
                });
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
                return json.value.map((item: any) => this._mapRoleAssignment(item));
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
        } catch (e) {
            console.error("Error fetching stats", e);
            return { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 };
        }
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

            if (groupsJson?.value) {
                const spGroups = groupsJson.value.map((g: any) => {
                    let users: any[] = [];
                    if (g.Users) {
                        users = Array.isArray(g.Users) ? g.Users : (g.Users.value || g.Users.results || []);
                    }
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
                        RoleAssignments: item.RoleAssignments ? item.RoleAssignments.map((ra: any) => this._mapRoleAssignment(ra)) : []
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
        // Strategy 1: /_api/site/usage
        const usage = await this._fetchUsage(`${this._webUrl}/_api/site/usage`, "Strategy 1");
        if (usage) return usage;

        // Strategy 2: /_api/site?$select=Usage
        const usage2 = await this._fetchUsage(`${this._webUrl}/_api/site?$select=Usage`, "Strategy 2");
        if (usage2) return usage2;

        return { storageUsed: 0, storageQuota: 0, usagePercentage: 0, lastItemModifiedDate: "" };
    }

    private async _fetchUsage(url: string, logPrefix: string): Promise<ISiteUsage | null> {
        try {
            const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (response.ok) {
                const json = await response.json();
                const usage = json.Usage || json.usage;
                if (usage) {
                    let quota = usage.StorageQuota || usage.storageQuota || 0;
                    const used = usage.Storage || usage.storage || 0;
                    const pct = usage.StoragePercentageUsed || usage.storagePercentageUsed || 0;

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
            console.warn(`${logPrefix} failed`, e);
        }
        return null;
    }

    private _mapRoleAssignment(item: any): IRoleAssignment {
        return {
            PrincipalId: item.PrincipalId,
            Member: item.Member,
            RoleDefinitionBindings: item.RoleDefinitionBindings
        };
    }

    public async getRoleDefinitions(): Promise<IRoleDefinitionDetail[]> { return []; }
    public async getUserEffectivePermissions(loginName: string): Promise<string[]> { return []; }
}
