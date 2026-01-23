import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup } from '../models/IPermissionData';

export interface IPermissionService {
    getSiteRoleAssignments(): Promise<IRoleAssignment[]>;
    getLists(excludedLists?: string[]): Promise<IListInfo[]>;
    getListRoleAssignments(listId: string, listTitle: string): Promise<IRoleAssignment[]>;
    getSiteStats(): Promise<ISiteStats>;
    getGroupMembers(groupId: number): Promise<IUser[]>;
    getUniquePermissionItems(listId: string): Promise<IItemPermission[]>;
    removeSitePermission(principalId: number): Promise<boolean>;
    removeListPermission(listId: string, principalId: number): Promise<boolean>;
    removeItemPermission(listId: string, itemId: number, principalId: number): Promise<boolean>;
    getSiteAdmins(): Promise<IUser[]>;
    getSiteOwners(): Promise<IUser[]>;
    getCurrentUser(): Promise<IUser>;
    searchUsers(query: string): Promise<IUser[]>;
    getSiteGroups(): Promise<IGroup[]>;
    getUserGroups(loginName: string): Promise<IGroup[]>;
    removeUserFromGroup(groupId: number, userId: number): Promise<boolean>;
}
