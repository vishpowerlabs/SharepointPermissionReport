export interface IUser {
    Id: number;
    Title: string;
    Email?: string;
    IsHiddenInUI: boolean;
    PrincipalType: number; // 1=User, 4=Security Group, 8=SharePoint Group
}

export interface IRoleDefinition {
    Id: number;
    Name: string;
    Description: string;
    Order: number;
    Hidden: boolean;
}

export interface IRoleAssignment {
    PrincipalId: number;
    Member: IUser;
    RoleDefinitionBindings: IRoleDefinition[];
}

export interface IListInfo {
    Id: string;
    Title: string;
    ItemCount: number;
    Hidden: boolean;
    ItemType: string; // List or Library
    ServerRelativeUrl: string;
    HasUniqueRoleAssignments: boolean;
    EntityTypeName: string;
}

export interface ISiteStats {
    totalUsers: number;
    totalGroups: number;
    uniquePermissionsCount: number;
}

export interface IItemPermission {
    Id: number;
    Title: string; // Or Name for files
    ServerRelativeUrl: string;
    FileSystemObjectType: number; // 0=File, 1=Folder
    RoleAssignments: IRoleAssignment[];
}
