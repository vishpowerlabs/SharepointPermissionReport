import { SPHttpClient, SPHttpClientResponse, MSGraphClientFactory } from '@microsoft/sp-http';
import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup, ISharingInfo, ISiteUsage, IRoleDefinitionDetail, IOversharedFolder } from '../models/IPermissionData';

export class PermissionService implements IPermissionService {
    private readonly _spHttpClient: SPHttpClient;
    private readonly _webUrl: string;
    private readonly _msGraphClientFactory?: MSGraphClientFactory;
    private readonly _userOrphanCache: Map<string, 'Deleted' | 'Disabled' | 'Active'> = new Map();
    private readonly _knownRemovedIds: Set<number> = new Set();


    constructor(spHttpClient: SPHttpClient, webUrl: string, msGraphClientFactory?: MSGraphClientFactory) {
        console.log(`[PermissionService] Constructor called. Instance created. WebUrl: ${webUrl}`);
        this._spHttpClient = spHttpClient;
        this._webUrl = webUrl;
        this._msGraphClientFactory = msGraphClientFactory;
    }

    public async getSiteRoleAssignments(): Promise<IRoleAssignment[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value.map((item: any) => this._mapRoleAssignment(item));
            }
            return [];
        } catch (error) {
            console.error("Error fetching site permissions", error);
            return [];
        }
    }

    public async getLists(excludedLists?: string[]): Promise<IListInfo[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/lists?$select=Id,Title,ItemCount,Hidden,BaseType,RootFolder/ServerRelativeUrl,HasUniqueRoleAssignments,EntityTypeName,BaseTemplate,LastItemModifiedDate,RootFolder/StorageMetrics/TotalSize&$filter=Hidden eq false&$expand=RootFolder,RootFolder/StorageMetrics`;
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
                        EntityTypeName: list.EntityTypeName,
                        TotalSize: list.RootFolder?.StorageMetrics?.TotalSize || 0,
                        LastItemModifiedDate: list.LastItemModifiedDate
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
            const endpoint = `${this._webUrl}/_api/web/lists/getbyid('${listId}')/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100`;
            const response = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
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
                uniquePermissionsCount: 0,
                emptyGroupsCount: groups.filter(g => g.UserCount === 0).length
            };
        } catch (error) {
            console.error("Error fetching stats", error);
            return { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0, emptyGroupsCount: 0 };
        }
    }

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
                    // Log the first group to inspect structure


                    let users: any[] = [];
                    if (g.Users) {
                        users = Array.isArray(g.Users) ? g.Users : (g.Users.value || g.Users.results || []);
                    }
                    return {
                        ...g,
                        PrincipalType: 8,
                        UserCount: users.length
                    };
                });
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
                    PrincipalType: 4,
                    UserCount: 1 // Assume 1 for AD group itself, as we can't expand members via this API easily
                }));
                allGroups = allGroups.concat(secGroups);
            }

            if (allGroups.length > 0) {


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



                return validGroups.map((g: any) => ({
                    Id: g.Id,
                    Title: g.Title,
                    LoginName: g.LoginName,
                    Description: g.Description || "",
                    IsHiddenInUI: g.IsHiddenInUI,
                    OwnerTitle: g.OwnerTitle || "",
                    PrincipalType: g.PrincipalType, // Map PrincipalType
                    UserCount: g.UserCount || 0
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

    public async getUniquePermissionItems(listId: string, signal?: AbortSignal): Promise<IItemPermission[]> {
        const results: IItemPermission[] = [];
        try {
            console.log(`Deep Scan: Starting scan for list ${listId} using GetItems (Recursive) + Batch Details`);

            // 1. Fetch Basic Items using standard REST (Flat, Recursive by default for IDs)
            // CRITICAL: Must use $orderby=ID asc to ensure we can paging past 5000 items (List View Threshold)
            // standard REST default sort might not use the ID index efficiently for large lists.
            let nextLink: string | undefined = `${this._webUrl}/_api/web/lists(guid'${listId}')/items?$select=ID,Title,FileRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments&$top=2000&$orderby=ID asc`;
            let hasMore = true;
            const candidates: any[] = [];

            while (hasMore && nextLink) {
                if (signal?.aborted) throw new Error('Scan aborted by user.');

                // Simple recursive query to get all items
                console.log(`Deep Scan: Fetching batch via GET: ${nextLink}`);
                const response: SPHttpClientResponse = await this._spHttpClient.get(
                    nextLink,
                    SPHttpClient.configurations.v1
                );

                if (!response.ok) {
                    console.error(`Deep Scan: Error fetching item IDs. Status: ${response.status}`);
                    break;
                }

                const json = await response.json();
                const items = json.value ? json.value : (json.d && json.d.results ? json.d.results : []);

                // Add to candidates
                candidates.push(...items);
                console.log(`Deep Scan: Fetched ${items.length} items. Total so far: ${candidates.length}`);

                // Update paging
                if (json['@odata.nextLink']) {
                    nextLink = json['@odata.nextLink'];
                } else if (json.d && json.d.__next) {
                    nextLink = json.d.__next;
                } else {
                    nextLink = undefined;
                    hasMore = false;
                }
            }

            console.log(`Deep Scan: Found ${candidates.length} candidate items (Total Fetched). Checking details for uniqueness...`);

            // 2. Batch Process Candidates to check Unique Permissions
            // We chunk candidates and use SPHttpClientBatch
            if (candidates.length > 0) {
                // Chunk size 40 (Batch limit is usually 100, keep safe)
                // Optimization: We already have HasUniqueRoleAssignments from REST call!
                // We only need to fetch Roles for those that are unique.
                // Filter batch to ONLY unique items logic?
                // Actually existing logic re-checks checks 'HasUniqueRoleAssignments'?
                // Let's optimize: checking HasUnique is redundant if we trust properties from fetch?
                // Yes, we selected HasUniqueRoleAssignments in the GET above.

                // Filter chunk to only Unique items to save batch calls?
                // Wait, Step 2 was "Check Unique Permissions".
                // Now that we have it from Step 1, we can skip directly to "Fetch Roles" (Step 3) for unique ones!

                // BUT, Step 3 requires 'HasUnique' check result. 
                // Let's rewrite this part. We can skip the 'Batch Check' step entirely and go straight to fetching roles for unique items.

                // Let's filter candidates down to unique ones immediatley.

                // REFACTOR: Skip Step 2 entirely if we have HasUniqueRoleAssignments
                const uniqueItems = candidates.filter(c => c.HasUniqueRoleAssignments === true);
                console.log(`Deep Scan: ${uniqueItems.length} items have unique permissions. Fetching roles...`);

                // 3. Fetch Roles for Unique Items
                if (uniqueItems.length > 0) {
                    // Reuse existing Step 3 logic but map 'r' structure correctly
                    const roleChunks = this.chunkArray(uniqueItems, 20); // Smaller chunks for heavy expansion
                    for (const rChunk of roleChunks) {
                        if (signal?.aborted) throw new Error('Scan aborted by user.');

                        // Parallel fetch for confirmed unique
                        await Promise.all(rChunk.map(async (item: any) => {
                            try {
                                const id = item.Id || item.ID;
                                const raEndpoint = `${this._webUrl}/_api/web/lists(guid'${listId}')/items(${id})/roleassignments?$expand=Member,RoleDefinitionBindings`;
                                const raResp = await this._spHttpClient.get(raEndpoint, SPHttpClient.configurations.v1);
                                if (raResp.ok) {
                                    const raJson = await raResp.json();
                                    const roleAssignments = raJson.value ? raJson.value.map((assignment: any) => this._mapRoleAssignment(assignment)) : [];

                                    // Map Item
                                    const fsObjType = (item.FileSystemObjectType === 1 || item.FSObjType == "1") ? 1 : 0;

                                    // Prioritize FileLeafRef from details if original missing
                                    const name = item.FileLeafRef || item.Title || `Item ${id}`;
                                    const url = item.FileRef;

                                    results.push({
                                        Id: Number.parseInt(id),
                                        Title: name,
                                        ServerRelativeUrl: url,
                                        FileSystemObjectType: fsObjType,
                                        RoleAssignments: roleAssignments
                                    });
                                }
                            } catch (e) {
                                console.warn(`Failed to fetch roles for item ${item.Id}`, e);
                            }
                        }));
                    }
                }

                // Return early to avoid running the old logic below
                return results;
            }

            // Old logic block below (needs to be removed or bypassed)
            // The original code had a `console.log` and `return results` here, which is now moved inside the `if (candidates.length > 0)` block.
            // This effectively removes the old logic.
            console.log(`Deep Scan: Finished. Total items scanned: ${candidates.length}. Total unique found: ${results.length}`);
            return results;

        } catch (error) {
            console.error(`Error scanning items for list ${listId}`, error);
            if (error instanceof Error && error.message === 'Scan aborted by user.') throw error;
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

    public async removeUserFromUserInfoList(principalId: number): Promise<boolean> {
        try {
            const endpoint = `${this._webUrl}/_api/web/siteusers/removebyid(${principalId})`;
            const response: SPHttpClientResponse = await this._spHttpClient.post(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                    }
                }
            );

            if (response.ok) {
                console.log(`[PermissionService] Successfully removed user ${principalId} from User Info List. Adding to mask.`);
                this._knownRemovedIds.add(principalId);
                return true;
            }

            // If user is not found, they are effectively removed
            if (response.status === 404) {
                console.warn(`User ${principalId} not found in User Info List, considering removed.`);
                this._knownRemovedIds.add(principalId);
                return true;
            }

            const errorText = await response.text();
            console.warn(`[PermissionService] Standard removebyid failed for user ${principalId}. Status: ${response.status}. Trying fallback deleteObject method...`);

            // Fallback: Try getting the user object and calling deleteObject (DELETE method)
            const fallbackEndpoint = `${this._webUrl}/_api/web/siteusers/getbyid(${principalId})`;
            const fallbackResponse: SPHttpClientResponse = await this._spHttpClient.post(
                fallbackEndpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'DELETE'
                    }
                }
            );

            if (fallbackResponse.ok) {
                console.log(`[PermissionService] Fallback deleteObject successful for user ${principalId}. Adding to mask.`);
                this._knownRemovedIds.add(principalId);
                return true;
            }

            const fallbackError = await fallbackResponse.text();
            console.error(`[PermissionService] Fallback delete also failed. Status: ${fallbackResponse.status}`, fallbackError);
            return false;

        } catch (error) {
            console.error(`Error removing user ${principalId} from User Info List`, error);
            return false;
        }
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

                const userId = item.EntityData?.SPUserID ? Number.parseInt(item.EntityData.SPUserID, 10) : -1;

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

    public async getExternalUsers(): Promise<IUser[]> {
        try {
            const allUsers = await this._getAllSiteUsers();
            // Filter for #ext# in LoginName or Email
            return allUsers.filter(u =>
                (u.LoginName?.toLowerCase().includes('#ext#')) ||
                (u.Email?.toLowerCase().includes('#ext#'))
            );
        } catch (error) {
            console.error("Error fetching external users", error);
            return [];
        }
    }

    public async getSharingLinks(): Promise<ISharingInfo[]> {
        try {
            // Strategy: Scan specific libraries (e.g. "Documents") for sharing links.
            // In a real generic tool, we might scan all libraries, but that is heavy.
            // We'll focus on the default "Documents" library for this "Report".

            const listTitle = "Documents";
            // Get items from Documents
            // We need the Absolute URL for GetObjectSharingInformation
            const itemsEndpoint = `${this._webUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title,FileRef,FileLeafRef,FileSystemObjectType&$top=20`;
            // Limiting to top 20 for performance in this demo. Real reporting would need paging/search.

            const response: SPHttpClientResponse = await this._spHttpClient.get(itemsEndpoint, SPHttpClient.configurations.v1);
            if (!response.ok) return [];

            const json = await response.json();
            const items = json.value || [];

            const sharingResults: ISharingInfo[] = [];

            // We must process sequentially or with limited concurrency to avoid throttling
            for (const item of items) {
                try {
                    // Construct absolute Object URL
                    // FileRef is server relative e.g. /sites/mysite/Shared Documents/foo.docx
                    // We need https://tenant.sharepoint.com/sites/mysite/...

                    // Simple hack: assume _webUrl doesn't have trailing slash? 
                    // But _webUrl might be https://tenant.sharepoint.com/sites/mysite
                    // We need the domain.

                    const location = globalThis.location;
                    const domain = `${location.protocol}//${location.hostname}`;
                    const objectUrl = `${domain}${item.FileRef}`;

                    const sharingInfoEndpoint = `${this._webUrl}/_api/SP.Sharing.ObjectSharingInformation.GetObjectSharingInformation`;

                    const payload = {
                        "request": {
                            "__metadata": { "type": "SP.Sharing.SharingInformationRequest" },
                            "objectUrl": objectUrl,
                            "retrieveAnonymousLinks": true,
                            "retrieveOrganizationSharingLinks": true, // Company-wide links
                            "retrieveSpecificPeopleSharingLinks": true,
                            "retrieveUserInfoDetails": true
                        }
                    };

                    const sharingResp = await this._spHttpClient.post(
                        sharingInfoEndpoint,
                        SPHttpClient.configurations.v1,
                        {
                            body: JSON.stringify(payload),
                            headers: {
                                'Accept': 'application/json;odata=verbose', // Verbose needed for some deep props? Try standard first.
                                'Content-Type': 'application/json;odata=verbose'
                            }
                        }
                    );

                    if (sharingResp.ok) {
                        const sharingJson = await sharingResp.json();
                        const data = sharingJson.d || sharingJson; // Handle verbose/noverbose

                        // Extract Anonymous Links
                        if (data.AnonymousEditLink || data.AnonymousViewLink) {
                            sharingResults.push({
                                documentName: item.FileLeafRef,
                                documentUrl: objectUrl,
                                linkType: "Anonymous (Anyone)",
                                sharedWith: ["Anyone with the link"]
                            });
                        }

                        // Extract specific sharing links?
                        // "SharingLinks" property contains list of links
                        if (data.SharingLinks?.results) {
                            data.SharingLinks.results.forEach((link: any) => {
                                // link.Url, link.IsActive, link.LinkKind
                                // LinkKind: 1=Organization, 2=Anonymous?, 3=SpecificPeople?
                                let type = "Specific People";
                                if (link.LinkKind === 1) type = "Organization (Internal)";
                                if (link.LinkKind === 2) type = "Anonymous";

                                // Avoid duplicate logging if we caught it above, but SharingLinks usually details specific generated links
                                sharingResults.push({
                                    documentName: item.FileLeafRef,
                                    documentUrl: objectUrl, // This is the item URL, not the *sharing* link URL. The user might want the sharing link itself?
                                    // link.Url is the SHARING link.
                                    // But we should put the sharing link in the report?
                                    // ISharingInfo has 'documentUrl', let's maybe put the Sharing Link in 'sharedWith' or add a field?
                                    // The user asked "provide the shareling".
                                    // I'll put the Sharing Link URL in 'documentUrl'??? No that's confusing.
                                    // I will append it to details.

                                    // Let's repurpose: documentUrl -> The Item URL.
                                    // Add details to sharedWith.

                                    sharedWith: [`Link: ${link.Url}`],
                                    linkType: type
                                });
                            });
                        }
                    }

                } catch (e) {
                    console.error("Error getting sharing info for " + item.FileLeafRef, e);
                }
            }

            return sharingResults;

        } catch (error) {
            console.error("Error fetching sharing links", error);
            // Fallback to empty if list not found
            return [];
        }
    }

    public async getOrphanedUsers(): Promise<IUser[]> {
        try {
            const allUsers = await this._getAllSiteUsers();

            // Filter potentially orphaned users
            // 1. Valid PrincipalType (User = 1)
            // 2. Not Site Admin (usually safely managed) - though admins can be orphaned, let's include them if they aren't "System"
            // 3. Not System Accounts

            const systemUsers = new Set(['System Account', 'SharePoint App', 'NT AUTHORITY\\authenticated users', 'spsearch']);

            const candidates = allUsers.filter(u =>
                u.PrincipalType === 1 &&
                !systemUsers.has(u.Title) &&
                !u.Title.includes('SharingLinks') &&
                !u.LoginName?.toLowerCase().includes('nt service') &&
                !u.LoginName?.toLowerCase().includes('sharepoint app')
            );
            if (candidates.length === 0 || !this._msGraphClientFactory) {
                return [];
            }

            return await this._checkOrphanCandidates(candidates);

        } catch (error) {
            console.error("Error fetching orphaned users", error);
            return [];
        }
    }

    public async debugItemPermissions(listTitle: string, itemQuery: string): Promise<string> {
        try {
            // 1. Get List
            const listUrl = `${this._webUrl}/_api/web/lists/getbytitle('${listTitle}')`;

            // 2. Search for item by FileLeafRef or Title
            // Using CAML for partial match
            const queryPayload = {
                query: {
                    ViewXml: `<View Scope="RecursiveAll">
                                    <Query>
                                        <Where>
                                            <Or>
                                                <Contains><FieldRef Name="FileLeafRef"/><Value Type="Text">${itemQuery}</Value></Contains>
                                                <Contains><FieldRef Name="Title"/><Value Type="Text">${itemQuery}</Value></Contains>
                                            </Or>
                                        </Where>
                                    </Query>
                                    <ViewFields>
                                        <FieldRef Name="ID"/><FieldRef Name="Title"/><FieldRef Name="FileLeafRef"/><FieldRef Name="FileRef"/><FieldRef Name="HasUniqueRoleAssignments"/>
                                    </ViewFields>
                                    <RowLimit>10</RowLimit>
                                  </View>`
                }
            };

            const endpoint = `${listUrl}/GetItems`;
            const response = await this._spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { body: JSON.stringify(queryPayload) });
            const json = await response.json();
            const items = json.value || (json.d && json.d.results) || [];

            if (items.length === 0) return `No items found matching '${itemQuery}' in '${listTitle}'.`;

            let log = `Found ${items.length} items. Analyzing first match:\n`;
            const item = items[0];
            const name = item.FileLeafRef || item.Title;

            log += `Item: ${name} (ID: ${item.ID || item.Id})\n`;
            log += `URL: ${item.FileRef}\n`;
            log += `HasUniqueRoleAssignments: ${item.HasUniqueRoleAssignments}\n`;

            if (!item.HasUniqueRoleAssignments) {
                log += `-> Inherits permissions from parent. Check parent permissions.\n`;
            } else {
                // Fetch Roles
                const roleUrl = `${listUrl}/items(${item.ID || item.Id})/roleassignments?$expand=Member,RoleDefinitionBindings`;
                const rResp = await this._spHttpClient.get(roleUrl, SPHttpClient.configurations.v1);
                const rJson = await rResp.json();
                const roles = rJson.value || (rJson.d && rJson.d.results) || [];

                log += `-> Found ${roles.length} role assignments:\n`;
                roles.forEach((r: any) => {
                    const type = r.Member.PrincipalType; // 1=User, 4=SecurityGroup, 8=SPGroup
                    log += `   - [${type}] ${r.Member.LoginName} (${r.Member.Title})\n`;
                    if (r.RoleDefinitionBindings) {
                        const perms = (r.RoleDefinitionBindings.results || r.RoleDefinitionBindings).map((d: any) => d.Name).join(', ');
                        log += `     Permissions: ${perms}\n`;
                    }
                });
            }

            return log;

        } catch (error: any) {
            return "Error during debug: " + (error.message || error);
        }
    }



    public async checkOrphansForRoleAssignments(roles: IRoleAssignment[]): Promise<IRoleAssignment[]> {
        try {
            const uniqueUsersMap = new Map<string, IUser>();

            roles.forEach(r => {
                if (r.Member.PrincipalType === 1) { // User
                    const key = r.Member.LoginName || r.Member.Email;
                    if (key && !uniqueUsersMap.has(key)) {
                        uniqueUsersMap.set(key, r.Member as IUser);
                    }
                }
            });

            if (uniqueUsersMap.size > 0 && this._msGraphClientFactory) {
                const usersToCheck = Array.from(uniqueUsersMap.values());
                const orphanedUsers = await this._checkOrphanCandidates(usersToCheck);

                // Create a map of orphan status for quick lookup (though objects are ref updated, this ensures we map back if needed)
                const orphanStatusMap = new Map<string, string>();
                orphanedUsers.forEach(u => {
                    const key = u.LoginName || u.Email;
                    if (key && u.OrphanStatus) {
                        orphanStatusMap.set(key, u.OrphanStatus);
                    }
                });

                // Ensure all roles reflect the status (in case of multiple roles for same user)
                roles.forEach(r => {
                    if (r.Member.PrincipalType === 1) {
                        const key = r.Member.LoginName || r.Member.Email;
                        if (key && orphanStatusMap.has(key)) {
                            r.Member.OrphanStatus = orphanStatusMap.get(key) as any;
                        }
                    }
                });
            }
            return roles;
        } catch (error) {
            console.error("Error checking orphans for roles", error);
            return roles;
        }
    }

    public maskOrphanedUser(userId: number): void {
        console.log(`[PermissionService] Masking orphaned user ID: ${userId}`);
        this._knownRemovedIds.add(userId);
    }

    public async checkOrphanUsers(users: IUser[]): Promise<IUser[]> {
        return this._checkOrphanCandidates(users);
    }

    private async _checkOrphanCandidates(candidates: IUser[]): Promise<IUser[]> {
        if (!this._msGraphClientFactory) return [];

        console.log("[OrphanCheck] Raw candidates:", candidates.map(c => ({ Title: c.Title, Type: c.PrincipalType, Id: c.Id })));

        // Filter out known removed IDs first AND ensure we only check Users (PrincipalType 1)
        // PrincipalType 4 are Security Groups, which will return 404 on /users endpoint
        // Also check if LoginName indicates a group claim (e.g. starts with c:0-.t)
        const effectiveCandidates = candidates.filter(u => {
            if (this._knownRemovedIds.has(u.Id)) return false;
            // Explicitly allow ONLY PrincipalType 1 (User)
            if (u.PrincipalType !== 1) return false;

            // Extra safety: Check LoginName patterns common for groups but sometimes mislabeled
            // c:0-.t, c:0+.w, c:0t.c are typically claims for groups/tenants/system
            // We'll exclude anything starting with c:0 to be safe as these aren't simple user object lookups
            if (u.LoginName && (u.LoginName.indexOf('c:0') === 0)) {
                return false;
            }

            return true;
        });

        if (effectiveCandidates.length === 0) return [];

        console.log(`[OrphanCheck] Total Candidates: ${candidates.length}, To Query Graph: ${effectiveCandidates.length}`);

        // Get Graph Client
        const graphClient = await this._msGraphClientFactory.getClient("3");

        const orphanedUsers: IUser[] = [];
        const toQuery: IUser[] = [];

        // 1. Check Cache
        for (const u of effectiveCandidates) {
            const key = u.Email || u.LoginName; // Simple key for now
            if (key && this._userOrphanCache.has(key)) {
                const status = this._userOrphanCache.get(key);
                if (status && status !== 'Active') {
                    u.OrphanStatus = status;
                    orphanedUsers.push(u);
                }
            } else {
                toQuery.push(u);
            }
        }

        console.log(`[OrphanCheck] Total Candidates: ${candidates.length}, To Query Graph: ${toQuery.length}`);

        if (toQuery.length === 0) {
            return orphanedUsers;
        }

        // 2. Query Uncached
        const batchSize = 20;
        const chunks = this.chunkArray(toQuery, batchSize);

        for (const chunk of chunks) {
            // Construct Batch Request
            const batchRequests = {
                requests: chunk.map((u, index) => {
                    let userIdentifier = u.Email;
                    if (!userIdentifier && u.LoginName) {
                        const parts = u.LoginName.split('|');
                        if (parts.length > 0) userIdentifier = parts[parts.length - 1];
                    }

                    return {
                        id: index.toString(),
                        method: "GET",
                        url: `/users/${userIdentifier}?$select=accountEnabled,userPrincipalName,displayName`
                    };
                })
            };

            try {
                // Send Batch
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
                const response = await graphClient.api('/$batch').post(batchRequests);

                // Process Responses
                if (response && response.responses) {
                    response.responses.forEach((res: any) => {
                        const reqId = Number(res.id);
                        const user = chunk[reqId];
                        const key = user.Email || user.LoginName;

                        if (res.status === 404) {
                            user.OrphanStatus = 'Deleted';
                            orphanedUsers.push(user);
                            if (key) this._userOrphanCache.set(key, 'Deleted');
                        } else if (res.status === 200) {
                            const adUser = res.body;
                            if (adUser.accountEnabled === false) {
                                user.OrphanStatus = 'Disabled';
                                orphanedUsers.push(user);
                                if (key) this._userOrphanCache.set(key, 'Disabled');
                            } else if (key) {
                                this._userOrphanCache.set(key, 'Active');
                            }
                        } else {
                            // Error - don't cache as active, might be transient
                        }
                    });
                }

                await new Promise(resolve => setTimeout(resolve, 200));

            } catch (batchError) {
                console.error("Batch request failed", batchError);
            }
        }
        console.log(`[OrphanCheck] Final Orphans Detected: ${orphanedUsers.map(u => u.LoginName).join(', ')}`);
        return orphanedUsers;
    }


    public async checkOversharingFolders(listId: string, signal?: AbortSignal): Promise<IOversharedFolder[]> {
        const results: IOversharedFolder[] = [];
        try {
            console.log(`[Oversharing] Starting SHALLOW folder scan for List ${listId}`);

            // 1. Get List Root Folder URL
            const listUrl = `${this._webUrl}/_api/web/lists(guid'${listId}')/rootfolder?$select=ServerRelativeUrl,Name`;
            const listResp = await this._spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            if (!listResp.ok) throw new Error("Failed to get list root folder");
            const listJson = await listResp.json();
            const rootUrl = listJson.ServerRelativeUrl;
            const encodedRootUrl = encodeURIComponent(rootUrl).replace(/'/g, "''");

            // 2. Scan Top-Level Folders Only
            const foldersUrl = `${this._webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedRootUrl}')/Folders?$select=Name,ServerRelativeUrl,ItemCount&$filter=Name ne 'Forms'&$top=2000`;
            const fResp = await this._spHttpClient.get(foldersUrl, SPHttpClient.configurations.v1);

            if (fResp && fResp.ok) {
                const fJson = await fResp.json();
                const folders = fJson.value || [];

                // Helper to fetch (reusing logic but simple fetch here is fine)
                const simpleFetcher = async (url: string) => this._spHttpClient.get(url, SPHttpClient.configurations.v1);

                for (const folder of folders) {
                    if (signal?.aborted) return results;
                    await this._checkItemForEveryone(folder.ServerRelativeUrl, true, results, simpleFetcher);
                }
            }

            return results;

        } catch (error: any) {
            console.error("Error checking folder oversharing", error);
            if (signal?.aborted) throw new Error("Scan aborted");
            return [];
        }
    }


    public async checkOversharingRootItems(listId: string, signal?: AbortSignal): Promise<IOversharedFolder[]> {
        const results: IOversharedFolder[] = [];
        try {
            console.log(`[Oversharing] Starting root item scan for List ${listId}`);

            // 1. Get List Root Folder URL
            const listUrl = `${this._webUrl}/_api/web/lists(guid'${listId}')/rootfolder?$select=ServerRelativeUrl,Name`;
            const listResp = await this._spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            if (!listResp.ok) throw new Error("Failed to get list root folder");
            const listJson = await listResp.json();
            const rootUrl = listJson.ServerRelativeUrl;
            const encodedRootUrl = encodeURIComponent(rootUrl).replace(/'/g, "''");

            // Retry Helper
            const fetchWithRetry = async (url: string, retries = 3): Promise<SPHttpClientResponse | null> => {
                for (let i = 0; i < retries; i++) {
                    if (signal?.aborted) return null;
                    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
                    if (response.status === 429 || response.status === 503) {
                        console.log(`[Oversharing] Throttled at ${url}, retrying...`);
                        await new Promise(resolve => setTimeout(resolve, 2000 * (i + 1)));
                        continue;
                    }
                    return response;
                }
                return null;
            };

            // 2. Scan Root Folders
            let skip = 0;
            const top = 5000;
            while (true) {
                if (signal?.aborted) break;
                const subUrl = `${this._webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedRootUrl}')/Folders?$select=Name,ServerRelativeUrl,ItemCount&$filter=Name ne 'Forms'&$top=${top}&$skip=${skip}`;
                const sResp = await fetchWithRetry(subUrl);
                if (!sResp || !sResp.ok) break;
                const sJson = await sResp.json();
                const folders = sJson.value || [];

                for (const folder of folders) {
                    if (signal?.aborted) return results;
                    await this._checkItemForEveryone(folder.ServerRelativeUrl, true, results, fetchWithRetry);
                }

                if (folders.length < top) break;
                skip += top;
            }

            // 3. Scan Root Files
            skip = 0;
            while (true) {
                if (signal?.aborted) break;
                const filesUrl = `${this._webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedRootUrl}')/Files?$select=Name,ServerRelativeUrl&$top=${top}&$skip=${skip}`;
                const fResp = await fetchWithRetry(filesUrl);
                if (!fResp || !fResp.ok) break;
                const fJson = await fResp.json();
                const files = fJson.value || [];

                for (const file of files) {
                    if (signal?.aborted) return results;
                    await this._checkItemForEveryone(file.ServerRelativeUrl, false, results, fetchWithRetry);
                }

                if (files.length < top) break;
                skip += top;
            }

            return results;

        } catch (error: any) {
            console.error("Error checking root items oversharing", error);
            if (signal?.aborted) throw new Error("Scan aborted");
            return [];
        }
    }

    private async _checkItemForEveryone(
        itemUrl: string,
        isFolder: boolean,
        results: IOversharedFolder[],
        fetcher: (url: string) => Promise<SPHttpClientResponse | null>
    ): Promise<boolean> {
        const encodedItemUrl = encodeURIComponent(itemUrl).replace(/'/g, "''");

        // Construct correct endpoint based on item type
        const typeSegment = isFolder ? `GetFolderByServerRelativeUrl('${encodedItemUrl}')` : `GetFileByServerRelativeUrl('${encodedItemUrl}')`;
        // Add timestamp to prevent caching of permission checks
        const rolesUrl = `${this._webUrl}/_api/web/${typeSegment}/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings&t=${new Date().getTime()}`;

        try {
            const rolesResp = await fetcher(rolesUrl);

            if (rolesResp && rolesResp.ok) {
                const rolesJson = await rolesResp.json();
                const roles = rolesJson.value || (rolesJson.d && rolesJson.d.results) || [];

                for (const r of roles) {
                    if (!r.Member) continue;

                    // Filter out Limited Access - strictly
                    const validBindings = (r.RoleDefinitionBindings || []).filter((d: any) => d.Name !== 'Limited Access');
                    if (validBindings.length === 0) continue; // User only has Limited Access, so skip

                    const title = r.Member.Title || '';
                    const login = r.Member.LoginName || '';

                    const isEveryone = (
                        title === 'Everyone' ||
                        title === 'Everyone except external users' ||
                        (login && login.indexOf('spo-grid-all-users') >= 0) ||
                        (login && login.indexOf('Everyone') >= 0) // Keeping broad for now but Limited Access filter should fix the main issue
                    );

                    if (isEveryone) {
                        const perms = validBindings.map((d: any) => d.Name).join(', ');
                        results.push({
                            Name: itemUrl.split('/').pop() || (isFolder ? 'Folder' : 'File'),
                            Path: itemUrl,
                            SharedWith: title,
                            LoginName: login,
                            Permissions: perms,
                            PrincipalType: r.Member.PrincipalType
                        });
                        console.log(`[Oversharing] MATCH! ${title} found at ${itemUrl}`);
                        return true;
                    }
                }
            }
        } catch (e) {
            console.warn(`[Oversharing] Error checking item ${itemUrl}`, e);
        }
        return false;
    }


    public async scanFolderContents(folderUrl: string, signal?: AbortSignal): Promise<IOversharedFolder[]> {
        console.log(`[Oversharing] Scanning folder contents for: ${folderUrl}`);
        const results: IOversharedFolder[] = [];

        // Helper for this scan to reuse Logic
        const fetchWithRetry = async (url: string, retries = 3): Promise<SPHttpClientResponse | null> => {
            let attempt = 0;
            while (attempt < retries) {
                if (signal?.aborted) return null;
                try {
                    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
                    if (response.status === 429) {
                        const retryAfter = response.headers.get('Retry-After');
                        const waitTime = retryAfter ? parseInt(retryAfter, 10) * 1000 : 2000 * Math.pow(2, attempt);
                        console.warn(`Throttled. Waiting ${waitTime}ms...`);
                        await new Promise(resolve => setTimeout(resolve, waitTime));
                        attempt++;
                        continue;
                    }
                    return response;
                } catch (err) {
                    console.error("Fetch error", err);
                    attempt++;
                }
            }
            return null;
        };

        try {
            // We want to scan the contents of THIS folder.
            // So we call _scanFoldersRecursive on this folder, BUT we need a way to force it to check children even if the Root itself is shared?
            // Actually, if the Root is shared, everything inside is shared.
            // But the user wants to see the items.
            // The existing _scanFoldersRecursive STOPS if it finds an Overshared folder.
            // We should arguably use a modified recursive function OR just reuse logic but ignore the "Stop" flag for the top level?

            // Let's implement a specific recursive scanner for this "Deep Dive" that records everything that is effectively shared.
            // If the parent is shared, should we list every single file inside? That could be huge.
            // Or should we just list unique permissions inside?
            // "Scan the folders and files" implies listing what is accessible.

            // Let's assume we want to find *explicitly* shared items inside, OR items that inherit from this shared folder?
            // If the folder itself is "Everyone", then listing 10,000 items inside as "Everyone" is spammy.
            // Maybe the user wants to check if *subfolders* have broken inheritance and are *more* open? Or just to verify?

            // Let's stick to the "Find Overshared Items" logic.
            // If I am scanning "Public", and it is shared with Everyone.
            // I want to see if "Public/Secret" is ALSO shared (inherited).
            // Yes, list it. The user wants to know what is exposed.
            // But maybe limit depth or count?

            // Let's use a simpler approach: Just scan immediate children (1 level deep? or recursive?)
            // "Scan folders and files" usually means everything.
            // Let's do recursive but be careful.

            await this._scanFolderContentsRecursive(folderUrl, results, signal, fetchWithRetry);

            return results;
        } catch (error) {
            console.error("Error scanning folder contents", error);
            return [];
        }
    }

    private async _scanFolderContentsRecursive(
        folderUrl: string,
        results: IOversharedFolder[],
        signal: AbortSignal | undefined,
        fetcher: (url: string) => Promise<SPHttpClientResponse | null>
    ) {
        if (signal?.aborted) return;
        const encodedFolderUrl = encodeURIComponent(folderUrl).replace(/'/g, "''");

        // 1. Get Folders
        const foldersUrl = `${this._webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedFolderUrl}')/Folders?$select=Name,ServerRelativeUrl,ItemCount&$filter=Name ne 'Forms'&$top=2000`;
        const fResp = await fetcher(foldersUrl);
        if (fResp && fResp.ok) {
            const fJson = await fResp.json();
            const folders = fJson.value || [];

            for (const folder of folders) {
                if (signal?.aborted) return;

                // Check permissions
                await this._checkItemForEveryone(folder.ServerRelativeUrl, true, results, fetcher);

                // Recurse
                await this._scanFolderContentsRecursive(folder.ServerRelativeUrl, results, signal, fetcher);
            }
        }

        // 2. Get Files
        const filesUrl = `${this._webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedFolderUrl}')/Files?$select=Name,ServerRelativeUrl&$top=2000`;
        const fileResp = await fetcher(filesUrl);
        if (fileResp && fileResp.ok) {
            const fileJson = await fileResp.json();
            const files = fileJson.value || [];

            for (const file of files) {
                if (signal?.aborted) return;
                await this._checkItemForEveryone(file.ServerRelativeUrl, false, results, fetcher);
            }
        }
    }
    private async _getAllSiteUsers(): Promise<IUser[]> {
        try {
            // Add timestamp to prevent browser caching
            const endpoint = `${this._webUrl}/_api/web/siteusers?t=${new Date().getTime()}`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

            if (!response.ok) {
                return [];
            }

            const json = await response.json();

            if (json?.value) {
                console.log(`[OrphanCheck] _getAllSiteUsers returned ${json.value.length} users from SharePoint.`);
                // Log known removed IDs to debug persistence
                console.log(`[OrphanCheck] _knownRemovedIds: ${Array.from(this._knownRemovedIds).join(', ')}`);

                const filtered = json.value.filter((u: any) => !this._knownRemovedIds.has(u.Id));
                console.log(`[OrphanCheck] After filtering known removed: ${filtered.length} users.`);

                return filtered.map((u: any) => ({
                    Id: u.Id,
                    Title: u.Title,
                    Email: u.Email,
                    LoginName: u.LoginName,
                    PrincipalType: u.PrincipalType,
                    IsHiddenInUI: u.IsHiddenInUI,
                    IsSiteAdmin: u.IsSiteAdmin
                } as IUser));
            }
            return [];
        } catch (error) {
            console.error("Error fetching all site users", error);
            return [];
        }
    }

    public async getSiteDetails(): Promise<{ Title: string; Url: string; HasUniqueRoleAssignments: boolean }> {
        try {
            const endpoint = `${this._webUrl}/_api/web?$select=Title,ServerRelativeUrl,HasUniqueRoleAssignments`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            return {
                Title: json.Title,
                Url: json.ServerRelativeUrl,
                HasUniqueRoleAssignments: json.HasUniqueRoleAssignments
            };
        } catch (error) {
            console.error("Error fetching site details", error);
            // Default safe fallback
            return { Title: "Current Site", Url: "", HasUniqueRoleAssignments: true };
        }
    }

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
            Member: {
                Id: item.Member.Id,
                Title: item.Member.Title,
                LoginName: item.Member.LoginName,
                Email: item.Member.Email || "",
                PrincipalType: item.Member.PrincipalType,
                IsHiddenInUI: item.Member.IsHiddenInUI || false
            },
            RoleDefinitionBindings: item.RoleDefinitionBindings ? item.RoleDefinitionBindings.filter((b: any) => b.Name !== 'Limited Access') : []
        };
    }


    public async getRoleDefinitions(): Promise<IRoleDefinitionDetail[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/roledefinitions`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            if (json?.value) {
                return json.value.map((r: any) => ({
                    Id: r.Id,
                    Name: r.Name,
                    Description: r.Description,
                    BasePermissions: r.BasePermissions,
                    Order: r.Order,
                    Hidden: r.Hidden,
                    RoleTypeKind: r.RoleTypeKind
                }));
            }
            return [];
        } catch (error) {
            console.error("Error fetching role definitions", error);
            return [];
        }
    }

    public async getUserEffectivePermissions(loginName: string): Promise<string[]> {
        try {
            // Need to encode login name safely
            const encodedLogin = encodeURIComponent(loginName);
            const endpoint = `${this._webUrl}/_api/web/getUserEffectivePermissions(@u)?@u='${encodedLogin}'`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const json = await response.json();

            // Returns { High: ..., Low: ... }
            const perms = json?.value || json;
            if (perms) {
                return this._parseBasePermissions(perms);
            }
            return [];

        } catch (error) {
            console.error("Error fetching effective permissions", error);
            return [];
        }
    }

    /**
     * Helper to convert BasePermissions (High/Low) into readable string array.
     * This is a simplified mapper. Full mapping takes ~30 definitions.
     */
    private readonly _permissionMapping = [
        { bit: 0x00000001, name: "View List Items" },
        { bit: 0x00000002, name: "Add List Items" },
        { bit: 0x00000004, name: "Edit List Items" },
        { bit: 0x00000008, name: "Delete List Items" },
        { bit: 0x00000010, name: "Approve Items" },
        { bit: 0x00000020, name: "Open Items" },
        { bit: 0x00000040, name: "View Versions" },
        { bit: 0x00000080, name: "Delete Versions" },
        { bit: 0x00000100, name: "Cancel Checkout" },
        { bit: 0x00000200, name: "Manage Personal Views" },
        { bit: 0x00000800, name: "Manage Lists" },
        { bit: 0x00001000, name: "View Pages" },
        { bit: 0x00002000, name: "Add and Customize Pages" },
        { bit: 0x00004000, name: "Apply Theme and Border" },
        { bit: 0x00008000, name: "Apply Style Sheets" },
        { bit: 0x00010000, name: "View Usage Data" },
        { bit: 0x00020000, name: "Create SSCSite" },
        { bit: 0x00040000, name: "Manage Subwebs" },
        { bit: 0x00080000, name: "Create Groups" },
        { bit: 0x00100000, name: "Manage Permissions" },
        { bit: 0x00200000, name: "Browse Directories" },
        { bit: 0x00400000, name: "Browse User Information" },
        { bit: 0x00800000, name: "Add Del Private Web Parts" },
        { bit: 0x01000000, name: "Update Personal Web Parts" },
        { bit: 0x02000000, name: "Manage Web" },
        { bit: 0x04000000, name: "Anonymous Search Access List" },
        { bit: 0x08000000, name: "Use Client Integration" },
        { bit: 0x10000000, name: "Use Remote APIs" },
        { bit: 0x20000000, name: "Manage Alerts" },
        { bit: 0x40000000, name: "Create Alerts" },
        { bit: 0x80000000, name: "Edit My User Information" },
    ];

    public async getAADGroupMembers(loginName: string, title: string): Promise<IUser[]> {
        if (!this._msGraphClientFactory) {
            console.warn("MSGraphClientFactory not available");
            return [];
        }

        try {
            const client = await this._msGraphClientFactory.getClient("3"); // Use v3 client

            // 1. Try to find the group in Graph
            // We search by displayName. This is not perfect but often sufficient for SP Groups mapping to AAD Groups.
            const groupsResponse = await client.api('/groups')
                .filter(`displayName eq '${title}'`)
                .select('id,displayName,groupTypes')
                .get();

            const groups = groupsResponse.value;
            if (!groups || groups.length === 0) {
                console.warn(`AAD Group '${title}' not found in Graph.`);
                return [];
            }

            // Use the first match
            const groupId = groups[0].id;

            // 2. Fetch members (transitive)
            const membersResponse = await client.api(`/groups/${groupId}/transitiveMembers`)
                .select('id,displayName,userPrincipalName,mail,accountEnabled')
                .top(100)
                .get();

            const members = membersResponse.value;

            // 3. Map to IUser
            return members.map((m: any) => ({
                Id: 0,
                Title: m.displayName,
                Email: m.mail || m.userPrincipalName,
                LoginName: m.userPrincipalName,
                PrincipalType: m['@odata.type'] === '#microsoft.graph.group' ? 8 : 1,
                IsHiddenInUI: false,
                IsSiteAdmin: false,
                OrphanStatus: m.accountEnabled === false ? 'Disabled' : undefined
            }));

        } catch (error) {
            console.error("Error fetching AAD group members", error);
            return [];
        }
    }

    private _parseBasePermissions(perms: { High: number; Low: number }): string[] {
        const rights: string[] = [];
        const low = perms.Low;
        const high = perms.High;

        this._permissionMapping.forEach(m => {
            if ((low & m.bit) > 0) rights.push(m.name);
        });

        if ((high & 0x00040000) > 0) {
            rights.push("Manage Web (Full Control)");
        }

        return rights;
    }
}
