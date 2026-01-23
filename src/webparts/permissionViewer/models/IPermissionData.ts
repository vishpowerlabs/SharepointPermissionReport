export interface IUser {
    Id: number;
    Title: string;
    Email?: string;
    LoginName?: string;
    IsHiddenInUI: boolean;
    PrincipalType: number;
    IsSiteAdmin?: boolean;
    IsSiteOwner?: boolean;
}

export interface IGroup {
    Id: number;
    Title: string;
    LoginName: string;
    Description: string;
    IsHiddenInUI: boolean;
    PrincipalType?: number; // 8 for SP Group, 4 for Security Group
    OwnerTitle: string;
    Users?: IUser[];
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
