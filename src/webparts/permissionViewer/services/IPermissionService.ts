import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup, ISharingInfo, ISiteUsage, IRoleDefinitionDetail, IOversharedFolder } from '../models/IPermissionData';

export interface IPermissionService {
    getSiteRoleAssignments(): Promise<IRoleAssignment[]>;
    getLists(excludedLists?: string[]): Promise<IListInfo[]>;
    getListRoleAssignments(listId: string, listTitle: string): Promise<IRoleAssignment[]>;
    getSiteStats(): Promise<ISiteStats>;
    getGroupMembers(groupId: number): Promise<IUser[]>;
    getUniquePermissionItems(listId: string, signal?: AbortSignal): Promise<IItemPermission[]>;
    removeSitePermission(principalId: number): Promise<boolean>;
    removeUserFromUserInfoList(principalId: number): Promise<boolean>;
    removeListPermission(listId: string, principalId: number): Promise<boolean>;
    removeItemPermission(listId: string, itemId: number, principalId: number): Promise<boolean>;
    getSiteAdmins(): Promise<IUser[]>;
    getSiteOwners(): Promise<IUser[]>;
    getCurrentUser(): Promise<IUser>;
    searchUsers(query: string): Promise<IUser[]>;
    getSiteGroups(): Promise<IGroup[]>;
    getUserGroups(loginName: string): Promise<IGroup[]>;
    removeUserFromGroup(groupId: number, userId: number): Promise<boolean>;
    getExternalUsers(): Promise<IUser[]>;

    getSharingLinks(): Promise<ISharingInfo[]>;
    debugItemPermissions(listTitle: string, itemQuery: string): Promise<string>; // Returns debug log // Sharing links often appear as users/groups
    getOrphanedUsers(): Promise<IUser[]>;
    getSiteDetails(): Promise<{ Title: string; Url: string; HasUniqueRoleAssignments: boolean }>;
    getSiteUsage(): Promise<ISiteUsage>;
    getRoleDefinitions(): Promise<IRoleDefinitionDetail[]>;
    getUserEffectivePermissions(loginName: string): Promise<string[]>;
    checkOrphansForRoleAssignments(roles: IRoleAssignment[]): Promise<IRoleAssignment[]>;
    maskOrphanedUser(userId: number): void;
    getAADGroupMembers(loginName: string, title: string): Promise<IUser[]>;
    checkOrphanUsers(users: IUser[]): Promise<IUser[]>;
    checkOversharingFolders(listId: string, signal?: AbortSignal): Promise<IOversharedFolder[]>;
    checkOversharingRootItems?(listId: string, signal?: AbortSignal): Promise<IOversharedFolder[]>;
    scanFolderContents?(folderUrl: string, signal?: AbortSignal): Promise<IOversharedFolder[]>;
}
