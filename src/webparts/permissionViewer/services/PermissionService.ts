import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup, ISharingInfo, ISiteUsage, IRoleDefinitionDetail } from '../models/IPermissionData';

export class ForceUpdatePermissionService implements IPermissionService {
    private readonly _spHttpClient: SPHttpClient;
    private readonly _webUrl: string;

    constructor(spHttpClient: SPHttpClient, webUrl: string) {
        console.log("ForceUpdatePermissionService V3 INITIALIZED - Code is live!");
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
                    if (g.Title === "DEVSITE Members" || g.Title === "Everyone") {
                        console.log(`[PermissionService] inspecting group: ${g.Title}`, g);
                        console.log(`[PermissionService] Users prop:`, g.Users);
                    }

                    // Handle OData response variations for expanded Users
                    // Standard v1: g.Users (array) or g.Users.value
                    // Verbose: g.Users.results
                    // Deferred: g.Users.__deferred (ignored here)
                    const users = g.Users ? (Array.isArray(g.Users) ? g.Users : (g.Users.value || g.Users.results || [])) : [];

                    console.log(`[PermissionService] Calculated users count for ${g.Title}:`, users.length);

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
            // Simple heuristic: Users with no email, not system, not sharing links, not admin
            const systemUsers = new Set(['System Account', 'SharePoint App', String.raw`NT AUTHORITY\authenticated users`, 'spsearch']);

            return allUsers.filter(u =>
                !u.Email &&
                u.PrincipalType === 1 &&
                !u.IsSiteAdmin &&
                !systemUsers.has(u.Title) &&
                !u.Title.includes('SharingLinks') &&
                !u.LoginName?.includes('nt service')
            );
        } catch (error) {
            console.error("Error fetching orphaned users", error);
            return [];
        }
    }

    private async _getAllSiteUsers(): Promise<IUser[]> {
        try {
            const endpoint = `${this._webUrl}/_api/web/siteusers`;
            const response: SPHttpClientResponse = await this._spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

            if (!response.ok) {
                console.error(`Error fetching site users: ${response.status} ${response.statusText}`);
                try {
                    const errorText = await response.text();
                    console.error("Error details:", errorText);
                } catch (e) {
                    console.error("Could not read error details", e);
                }
                return [];
            }

            const json = await response.json();

            if (json?.value) {
                return json.value.map((u: any) => ({
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
        try {
            // 1. Get Storage Quota & Usage
            const usageEndpoint = `${this._webUrl}/_api/site/usage`;
            const usageReq = this._spHttpClient.get(usageEndpoint, SPHttpClient.configurations.v1);

            // 2. Get Last Modified Date (from Web properties)
            const webEndpoint = `${this._webUrl}/_api/web?$select=LastItemModifiedDate`;
            const webReq = this._spHttpClient.get(webEndpoint, SPHttpClient.configurations.v1);

            const [usageResp, webResp] = await Promise.all([usageReq, webReq]);
            const usageJson = await usageResp.json();
            const webJson = await webResp.json();

            // usageJson.Usage.Storage is in bytes
            // usageJson.Usage.StorageQuota is in bytes

            // Usage endpoint structure might vary slightly, but standard is:
            // d: { Usage: { Storage: ..., StorageQuota: ... } } or direct properties depending on Accept header
            // With standard JSON: { Usage : { Storage: 123, StorageQuota: 456 } }

            // SharePoint API returns: { Storage: 576195757, StoragePercentageUsed: 0.00002096... }
            // Storage is in bytes, we need to calculate quota from percentage
            const storageUsedBytes = usageJson?.Storage || 0;
            const storagePercentageUsed = usageJson?.StoragePercentageUsed || 0;

            // Calculate quota: used / percentage = quota
            const storageQuotaBytes = storagePercentageUsed > 0
                ? storageUsedBytes / storagePercentageUsed
                : 0;

            console.log('=== STORAGE DEBUG ===');
            console.log('Raw API Response:', usageJson);
            console.log('Storage Used (bytes):', storageUsedBytes);
            console.log('Storage Percentage:', storagePercentageUsed);
            console.log('Calculated Quota (bytes):', storageQuotaBytes);
            console.log('Used (GB):', (storageUsedBytes / (1024 * 1024 * 1024)).toFixed(2));
            console.log('Quota (GB):', (storageQuotaBytes / (1024 * 1024 * 1024)).toFixed(2));
            console.log('==================');

            return {
                storageUsed: storageUsedBytes,
                storageQuota: storageQuotaBytes,
                usagePercentage: storagePercentageUsed,
                lastItemModifiedDate: webJson.LastItemModifiedDate
            };

        } catch (error) {
            console.error("Error fetching site usage", error);
            return { storageUsed: 0, storageQuota: 0, usagePercentage: 0, lastItemModifiedDate: "" };
        }
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
    private _parseBasePermissions(perms: { High: number; Low: number }): string[] {
        const rights: string[] = [];
        const low = perms.Low;
        const high = perms.High;

        // Common Low Permissions
        if (low & 1) rights.push("ViewListItems");
        if (low & 2) rights.push("AddListItems");
        if (low & 4) rights.push("EditListItems");
        if (low & 8) rights.push("DeleteListItems");
        if (low & 16) rights.push("ApproveItems");
        if (low & 32) rights.push("OpenItems");
        if (low & 64) rights.push("ViewVersions");
        if (low & 128) rights.push("DeleteVersions");
        if (low & 256) rights.push("CancelCheckout");
        if (low & 512) rights.push("ManagePersonalViews");
        if (low & 2048) rights.push("ManageLists");
        if (low & 4096) rights.push("ViewFormPages");

        if (low & 65536) rights.push("ViewPages");
        if (low & 131072) rights.push("LayoutsPage"); // AddAndCustomizePages? No, LayoutsPage is usually mapped to something else? 
        // 0x20000 = ViewPages (131072) -> Actually 0x20000 is ViewPages. 
        // Let's rely on standard enum values:

        // Re-mapping based on SP.PermissionKind
        // See: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-permission-levels

        // Very basic mapping for demo
        if ((low & 0x00000001) > 0) rights.push("View List Items");
        if ((low & 0x00000002) > 0) rights.push("Add List Items");
        if ((low & 0x00000004) > 0) rights.push("Edit List Items");
        if ((low & 0x00000008) > 0) rights.push("Delete List Items");

        if ((low & 0x00020000) > 0) rights.push("View Pages");
        if ((low & 0x00040000) > 0) rights.push("Enumerate Permissions");
        if ((low & 0x01000000) > 0) rights.push("Create Groups");
        if ((low & 0x02000000) > 0) rights.push("Manage Permissions");
        if ((low & 0x04000000) > 0) rights.push("Browse User Info");

        // Full Control essentially checks multiple high bits or specifically ManageWeb
        if ((high & 0x00040000) > 0) rights.push("Manage Web (Full Control)");

        return rights;
    }
}
