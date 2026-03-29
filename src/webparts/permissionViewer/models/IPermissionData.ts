export interface IUser {
    Id: number;
    Title: string;
    Email?: string;
    LoginName?: string;
    IsHiddenInUI: boolean;
    PrincipalType: number;
    IsSiteAdmin?: boolean;
    IsSiteOwner?: boolean;
    OrphanStatus?: 'Deleted' | 'Disabled' | 'Active';
    UserPrincipalName?: string;
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
    UserCount: number; // Mandatory now to avoid displaying issues
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
    TotalSize?: number;
    LastItemModifiedDate?: string;
}

export interface ISiteStats {
    totalUsers: number;
    totalGroups: number;
    uniquePermissionsCount: number;
    emptyGroupsCount: number;
}

export interface IItemPermission {
    Id: number;
    Title: string; // Or Name for files
    ServerRelativeUrl: string;
    FileSystemObjectType: number; // 0=File, 1=Folder
    RoleAssignments: IRoleAssignment[];
}

export interface ISharingInfo {
    documentName: string;
    documentUrl: string;
    sharedWith: string[]; // e.g. ["someone@gmail.com", "Everyone"]
    linkType: string; // e.g. "Anonymous", "Organization", "Specific People"
}

export interface ISiteUsage {
    storageUsed: number;
    storageQuota: number;
    usagePercentage: number;
    lastItemModifiedDate: string;
}

export interface IRoleDefinitionDetail {
    Id: number;
    Name: string;
    Description: string;
    BasePermissions: { High: number; Low: number; };
    Order: number;
    Hidden: boolean;
    RoleTypeKind: number;
}



export interface ICommonProps {
    webPartTitle: string;
    // ... other props
    webPartTitleFontSize?: number;
    headerOpacity?: number;
    themeVariant?: any;
    storageFormat?: 'MB' | 'GB' | 'Auto';
    contentFontSize?: string;
    showComponentHeader?: boolean;
    navLayout?: 'left' | 'top';

    // Mock data
    useMockData?: boolean;
    simulateAccessDenied?: boolean;

    // Config
    excludedLists?: string[];
    showSecurityGovernanceTab?: boolean;
    customCss?: string;
    showStats?: boolean;
    buttonFontSize?: string;
    showExternalUserAudit?: boolean;
    showSharingLinks?: boolean;
    showOrphanedUsers?: boolean;
}

export interface IPermissionViewerWebPartProps extends ICommonProps {
    description: string;
}

export interface IPublicAccessResult {
    key: string;
    listName: string;
    listId: string;
    itemId?: number;
    itemName: string;
    itemUrl: string;
    scope: 'List' | 'Item' | 'Folder';
    principalName: string;
    principalType: string; // 'Everyone', 'Anonymous', 'Authenticated'
    roles: string[];
}


export interface IOversharedFolder {
    Name: string;
    Path: string;
    SharedWith: string;
    LoginName?: string;
    Permissions: string;
    PrincipalType?: number;
}
