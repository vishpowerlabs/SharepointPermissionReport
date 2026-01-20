import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission } from '../models/IPermissionData';

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
}
