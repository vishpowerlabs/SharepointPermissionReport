import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup } from '../models/IPermissionData';

export class PermissionServiceImpl implements IPermissionService {
    private readonly _spHttpClient: SPHttpClient;
    private readonly _webUrl: string;

    constructor(spHttpClient: SPHttpClient, webUrl: string) {
        this._spHttpClient = spHttpClient;
        this._webUrl = webUrl;
    }

    public async getSiteRoleAssignments(): Promise<IRoleAssignment[]> {
        try {
            // Expand Member to get user details, and RoleDefinitionBindings to get permission levels
            const endpoint = `${this._webUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return this._processRoleAssignments(json.value as IRoleAssignment[]);
            }
            return [];
        } catch (error) {
            console.error("Error fetching site permissions", error);
            return [];
        }
    }

    public async getLists(excludedLists?: string[]): Promise<IListInfo[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/lists?$select=Id,Title,ItemCount,Hidden,BaseType,RootFolder/ServerRelativeUrl,HasUniqueRoleAssignments,EntityTypeName,BaseTemplate&$filter=Hidden eq false&$expand=RootFolder`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                // Filter out system lists based on requirements
                const systemLists = new Set(excludedLists || [
                    "Site Assets", "Style Library", "Master Page Gallery",
                    "Form Templates", "User Information List", "Composed Looks", "Solution Gallery",
                    "TaxonomyHiddenList", "Appdata", "Appfiles"
                ]);

                return json.value.filter((list: any) => {
                    // Exclude catalogs (BaseTemplate 113, 116 etc usually, but we check names/types)
                    // BaseType 0 = Generic List, 1 = Document Library

                    if (systemLists.has(list.Title)) return false;
                    if (list.EntityTypeName === "OData__catalogs") return false;

                    return true;
                }).map((list: any) => {
                    return {
                        Id: list.Id,
                        Title: list.Title,
                        ItemCount: list.ItemCount,
                        Hidden: list.Hidden,
                        ItemType: list.BaseType === 1 ? 'Library' : 'List',
                        ServerRelativeUrl: list.RootFolder ? list.RootFolder.ServerRelativeUrl : '',
                        HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
                        EntityTypeName: list.EntityTypeName
                    } as IListInfo;
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
            // We can use GetByTitle or GetById. ById is safer.
            const endpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return this._processRoleAssignments(json.value as IRoleAssignment[]);
            }
            return [];
        } catch (error) {
            console.error(`Error fetching permissions for list ${listTitle}`, error);
            return [];
        }
    }

    public async getSiteStats(): Promise<ISiteStats> {
        try {
            // Fetch Site Users
            const usersEndpoint = `${this._webUrl}/_api/web/siteusers`;
            const usersReq = this._spHttpClient.get(usersEndpoint, SPHttpClient.configurations.v1);

            // Fetch Groups using the new method to ensure consistency
            const groupsReq = this.getSiteGroups();

            const [usersResp, groups] = await Promise.all([usersReq, groupsReq]);
            const usersJson = await usersResp.json();

            return {
                totalUsers: usersJson.value ? usersJson.value.length : 0,
                totalGroups: groups.length,
                uniquePermissionsCount: 0
            };
        } catch (error) {
            console.error("Error fetching stats", error);
            return { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 };
        }
    }

    public async getSiteGroups(): Promise<IGroup[]> {
        try {
            // 1. Fetch SharePoint Groups
            const groupsEndpoint = `${this._webUrl}/_api/web/sitegroups`;
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
                const spGroups = groupsJson.value.map((g: any) => ({ ...g, PrincipalType: 8 })); // Add Type 8 explicitly if needed
                allGroups = allGroups.concat(spGroups);
            }

            // Process Security Groups
            if (secGroupsJson?.value) {
                // Security groups from siteusers don't have "OwnerTitle", but have "Title", "LoginName", "Id"
                const secGroups = secGroupsJson.value.map((g: any) => ({
                    Id: g.Id,
                    Title: g.Title,
                    LoginName: g.LoginName,
                    Description: g.Email || "", // Use Email as description for security groups typically
                    OwnerTitle: "Active Directory", // Static owner for AD groups
                    IsHiddenInUI: g.IsHiddenInUI,
                    PrincipalType: 4
                }));
                allGroups = allGroups.concat(secGroups);
            }

            if (allGroups.length > 0) {
                console.log("Raw All Groups:", allGroups.map((g: any) => g.Title));

                // Filter out hidden groups and specific system groups
                const validGroups = allGroups.filter((g: any) => {
                    if (g.IsHiddenInUI) return false;

                    // Exclude specific system group names
                    const systemGroupNames = [
                        "Excel Services Viewers",
                        "Restricted Readers",
                        "Translation Managers",
                        "Designers",
                        "Approvers",
                        "Hierarchy Managers",
                        "Quick Deploy Users"
                    ];

                    const title = g.Title || "";
                    const login = g.LoginName || "";

                    // Check for SharingLinks (case insensitive)
                    if (login.toLowerCase().includes("sharinglinks")) return false;
                    if (title.toLowerCase().includes("sharinglinks")) return false;

                    // Check exact matches for known system groups
                    if (systemGroupNames.includes(title)) return false;

                    return true;
                });

                console.log("Filtered Groups:", validGroups.map((g: any) => g.Title));

                return validGroups.map((g: any) => ({
                    Id: g.Id,
                    Title: g.Title,
                    LoginName: g.LoginName,
                    Description: g.Description || "",
                    IsHiddenInUI: g.IsHiddenInUI,
                    OwnerTitle: g.OwnerTitle || "",
                    PrincipalType: g.PrincipalType // Map PrincipalType
                } as IGroup));
            }
            return [];
        } catch (error) {
            console.error("Error fetching site groups", error);
            return [];
        }
    }

    public async getGroupMembers(groupId: number): Promise<IUser[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/sitegroups/getbyid(${groupId})/users`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value as IUser[];
            }
            return [];
        } catch (error) {
            console.error(`Error fetching members for group ${groupId}`, error);
            return [];
        }
    }

    public async getUniquePermissionItems(listId: string): Promise<IItemPermission[]> {
        try {
            // HasUniqueRoleAssignments is often not filterable in OData on items endpoint.
            // Strategy: 
            // 1. Fetch all items (up to 5000) with minimal fields including HasUniqueRoleAssignments.
            // 2. Filter client-side.
            // 3. For each unique item, fetch role assignments.

            // Note: To ensure we get all items (recursive), we rely on 'items' returning them.
            // Ideally we'd use CAML with Scope='RecursiveAll', but standard items endpoint usually returns items flat or based on default view.
            // We'll trust 'items' for now but handle the filtering client side.

            const endpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/items?$select=Id,Title,FileRef,FileLeafRef,FileSystemObjectType,HasUniqueRoleAssignments&$top=5000`;

            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                const allItems: any[] = json.value; // Keep any here for raw JSON but cast specifically later if needed or define interface
                const uniqueItems = allItems.filter((i: { HasUniqueRoleAssignments: boolean }) => i.HasUniqueRoleAssignments === true);

                if (uniqueItems.length === 0) return [];

                // Now fetch permissions for these specific items
                // To avoid making too many parallel calls, we'll process them in chunks or Promise.all

                const results: IItemPermission[] = [];

                // Limit parallelism to avoid 429 throttling
                const chunks = this.chunkArray(uniqueItems, 5);

                for (const chunk of chunks) {
                    const promises = chunk.map(async (item: any) => {
                        try {
                            // Fetch roles
                            const permEndpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/items(${item.Id})/roleassignments?$expand=Member,RoleDefinitionBindings`;
                            const permResp = await this._spHttpClient.get(permEndpoint, SPHttpClient.configurations.v1);
                            const permJson = await permResp.json();
                            const roles = permJson?.value ? this._processRoleAssignments(permJson.value as IRoleAssignment[]) : [];

                            return {
                                Id: item.Id,
                                Title: item.FileLeafRef || item.Title, // Use filename for files
                                ServerRelativeUrl: item.FileRef,
                                FileSystemObjectType: item.FileSystemObjectType,
                                RoleAssignments: roles
                            } as IItemPermission;
                        } catch (e) {
                            console.error(`Error fetching perms for item ${item.Id}`, e);
                            return null;
                        }
                    });

                    const chunkResults = await Promise.all(promises);
                    chunkResults.forEach((r) => { if (r) results.push(r); });
                }

                return results;
            }
            return [];
        } catch (error) {
            console.error(`Error scanning items for list ${listId}`, error);
            return [];
        }
    }

    private chunkArray<T>(myArray: T[], chunk_size: number): T[][] {
        const results: T[][] = [];
        const arrayCopy = [...myArray];
        while (arrayCopy.length) {
            results.push(arrayCopy.splice(0, chunk_size));
        }
        return results;
    }
    public async removeSitePermission(principalId: number): Promise<boolean> {
        const endpoint = `${this._webUrl}/_api/web/roleassignments/getbyprincipalid(${principalId})`;
        return this._removeRoleAssignment(endpoint);
    }

    public async removeListPermission(listId: string, principalId: number): Promise<boolean> {
        const endpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/roleassignments/getbyprincipalid(${principalId})`;
        return this._removeRoleAssignment(endpoint);
    }

    public async removeItemPermission(listId: string, itemId: number, principalId: number): Promise<boolean> {
        const endpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/roleassignments/getbyprincipalid(${principalId})`;
        return this._removeRoleAssignment(endpoint);
    }

    private async _removeRoleAssignment(endpoint: string): Promise<boolean> {
        try {
            const response: SPHttpClientResponse = await this._spHttpClient.post(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'DELETE',
                        'IF-MATCH': '*'
                    }
                }
            );
            return response.ok;
        } catch (error) {
            console.error("Error removing permission", error);
            return false;
        }
    }

    private _processRoleAssignments(assignments: IRoleAssignment[]): IRoleAssignment[] {
        if (!assignments) return [];

        return assignments.map(assignment => {
            // Filter out "Limited Access" from bindings
            const filteredBindings = assignment.RoleDefinitionBindings.filter(binding => binding.Name !== 'Limited Access');

            return {
                ...assignment,
                RoleDefinitionBindings: filteredBindings
            };
        }).filter(assignment => assignment.RoleDefinitionBindings.length > 0);
    }
    public async searchUsers(query: string): Promise<IUser[]> {
        try {
            // Use ClientPeoplePickerSearchUser to search the entire directory (AD/AAD)
            const endpoint = `${this._webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
            const payload = {
                queryParams: {
                    QueryString: query,
                    MaximumEntitySuggestions: 10,
                    AllowEmailAddresses: true,
                    AllowOnlyEmailAddresses: false,
                    PrincipalType: 15, // Users (1), Distribution Lists (2), Security Groups (4), SharePoint Groups (8) => 1+2+4+8 = 15
                    PrincipalSource: 15, // All sources
                    SharePointGroupID: 0
                }
            };

            const response: SPHttpClientResponse = await this._spHttpClient.post(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    body: JSON.stringify(payload),
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                    }
                }
            );

            const json = await response.json();

            // Parse result (people picker returns a serialized JSON string in the 'value' property sometimes, or directly arrays)
            // The structure typically is: json.value = "[{...}, {...}]" (stringified JSON)
            const resultsData = json?.value ? JSON.parse(json.value) : [];

            // Map to IUser interface
            return resultsData.map((item: any) => {
                // Ensure we get a valid ID. People Picker results might not have the site User ID if they haven't visited yet.
                // However, for purposes of "Check Access", we often need their ID to check permissions.
                // If they have no permissions, they might not be in the User Info List, so no ID.
                // But the caller (Check Access) expects an ID for `getSiteAccess` and `runScan` which rely on `Member.Id`.

                // CRITICAL LIMITATION: If we search for a fresh user who is NOT in the User Info List, we cannot check their permissions easily via REST API 
                // because all permission endpoints require a Principal ID (int). 
                // We can't resolve permissions for a user who "doesn't exist" in the site.

                // For the purpose of this tool ("Check Access"), if they aren't in the site, they likely have NO permissions (unless via AD group).
                // So, we should try to return them, but maybe flag them or handle the missing ID gracefully in the UI.

                // Actually, let's use `EntityData.SPUserID` if available, or try `ensureUser` if we were really robust, 
                // but `ensureUser` is a write operation (adds them to site), which might be aggressive for a read-only report.

                // Let's rely on what we get. If EntityData.SPUserID is present, use it.
                // If not, use -1 or similar to indicate "Not in site".

                const userId = item.EntityData && item.EntityData.SPUserID ? parseInt(item.EntityData.SPUserID, 10) : -1;

                return {
                    Id: userId,
                    Title: item.DisplayText,
                    Email: item.EntityData?.Email || item.Description || "", // Email is often in EntityData or Description
                    LoginName: item.Key,
                    PrincipalType: item.EntityType === "User" ? 1 : 4, // Simplified mapping
                    IsHiddenInUI: false
                } as IUser;
            });

        } catch (error) {
            console.error("Error searching users", error);
            return [];
        }
    }

    public async getSiteAdmins(): Promise<IUser[]> {
        try {
            // Filter for users where IsSiteAdmin is true
            const endpoint = `${this._webUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value.map((u: any) => ({
                    Id: u.Id,
                    Title: u.Title,
                    Email: u.Email,
                    LoginName: u.LoginName,
                    PrincipalType: u.PrincipalType,
                    IsHiddenInUI: u.IsHiddenInUI,
                    IsSiteAdmin: true
                } as IUser));
            }
            return [];
        } catch (error) {
            console.error("Error fetching site admins", error);
            return [];
        }
    }
    public async getSiteOwners(): Promise<IUser[]> {
        try {
            // AssociatedOwnerGroup might be null if not set, handled by catch or empty check
            const endpoint = `${this._webUrl}/_api/web/associatedownergroup/users`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value.map((u: any) => ({
                    Id: u.Id,
                    Title: u.Title,
                    Email: u.Email,
                    LoginName: u.LoginName,
                    PrincipalType: u.PrincipalType,
                    IsHiddenInUI: u.IsHiddenInUI,
                    IsSiteOwner: true
                } as IUser));
            }
            return [];
        } catch (error) {
            console.error("Error fetching site owners", error);
            return [];
        }
    }

    public async getCurrentUser(): Promise<IUser> {
        try {
            const endpoint = `${this._webUrl}/_api/web/currentuser`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const u = await response.json();

            return {
                Id: u.Id,
                Title: u.Title,
                Email: u.Email,
                LoginName: u.LoginName,
                PrincipalType: u.PrincipalType,
                IsSiteAdmin: u.IsSiteAdmin,
                IsHiddenInUI: u.IsHiddenInUI
            } as IUser;
        } catch (error) {
            console.error("Error fetching current user", error);
            throw error;
        }
    }
    public async getUserGroups(loginName: string): Promise<IGroup[]> {
        try {
            const encodedLogin = encodeURIComponent(loginName);
            const endpoint = `${this._webUrl}/_api/web/siteusers(@u)/groups?@u='${encodedLogin}'`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value.map((g: any) => ({
                    Id: g.Id,
                    Title: g.Title,
                    LoginName: g.LoginName,
                    Description: g.Description || "",
                    PrincipalType: g.PrincipalType,
                    IsHiddenInUI: g.IsHiddenInUI,
                    OwnerTitle: g.OwnerTitle || ""
                } as IGroup));
            }
            return [];
        } catch (error) {
            console.error(`Error fetching groups for user ${loginName}`, error);
            return [];
        }
    }

    public async removeUserFromGroup(groupId: number, userId: number): Promise<boolean> {
        try {
            const endpoint = `${this._webUrl}/_api/web/sitegroups/getbyid(${groupId})/users/removebyid(${userId})`;
            const response: SPHttpClientResponse = await this._spHttpClient.post(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'DELETE'
                    }
                }
            );
            return response.ok;
        } catch (error) {
            console.error(`Error removing user ${userId} from group ${groupId}`, error);
            return false;
        }
    }
}
