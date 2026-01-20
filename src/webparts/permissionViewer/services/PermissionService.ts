import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission } from '../models/IPermissionData';

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
            const groupsEndpoint = `${this._webUrl}/_api/web/sitegroups`;

            const usersReq = this._spHttpClient.get(usersEndpoint, SPHttpClient.configurations.v1);
            const groupsReq = this._spHttpClient.get(groupsEndpoint, SPHttpClient.configurations.v1);

            const [usersResp, groupsResp] = await Promise.all([usersReq, groupsReq]);

            const usersJson = await usersResp.json();
            const groupsJson = await groupsResp.json();

            return {
                totalUsers: usersJson.value ? usersJson.value.length : 0,
                totalGroups: groupsJson.value ? groupsJson.value.length : 0,
                uniquePermissionsCount: 0 // To be calculated by caller or separate logic
            };
        } catch (error) {
            console.error("Error fetching stats", error);
            return { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 };
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
}
