import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission } from '../models/IPermissionData';

export interface IPermissionService {
    getSiteRoleAssignments(): Promise<IRoleAssignment[]>;
    getLists(excludedLists?: string[]): Promise<IListInfo[]>;
    getListRoleAssignments(listId: string, listTitle: string): Promise<IRoleAssignment[]>;
    getSiteStats(): Promise<ISiteStats>;
    getGroupMembers(groupId: number): Promise<IUser[]>;
    getUniquePermissionItems(listId: string): Promise<IItemPermission[]>;
}
