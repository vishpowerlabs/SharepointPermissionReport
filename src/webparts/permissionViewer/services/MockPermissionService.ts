import { IPermissionService } from './IPermissionService';
import { IRoleAssignment, IListInfo, ISiteStats, IUser, IItemPermission, IGroup, ISharingInfo, ISiteUsage, IRoleDefinitionDetail } from '../models/IPermissionData';

export class MockPermissionService implements IPermissionService {

    public async getSiteRoleAssignments(): Promise<IRoleAssignment[]> {
        return [
            {
                Member: { Id: 1, Title: "Marketing Owners", LoginName: "c:0+.w|s-1-5-21...", PrincipalType: 8, IsHiddenInUI: false },
                PrincipalId: 1,
                RoleDefinitionBindings: [{ Id: 1073741829, Name: "Full Control", Description: "Has full control.", Order: 1, Hidden: false }]
            },
            {
                Member: { Id: 2, Title: "Marketing Members", LoginName: "c:0+.w|s-1-5-21...", PrincipalType: 8, IsHiddenInUI: false },
                PrincipalId: 2,
                RoleDefinitionBindings: [{ Id: 1073741830, Name: "Edit", Description: "Can add, edit and delete lists.", Order: 2, Hidden: false }]
            },
            {
                Member: { Id: 3, Title: "Marketing Visitors", LoginName: "c:0+.w|s-1-5-21...", PrincipalType: 8, IsHiddenInUI: false },
                PrincipalId: 3,
                RoleDefinitionBindings: [{ Id: 1073741826, Name: "Read", Description: "Can view pages and list items.", Order: 3, Hidden: false }]
            },
            {
                Member: { Id: 4, Title: "Sarah Jenkins", LoginName: "i:0#.f|membership|sarah.j@contoso.com", Email: "sarah.j@contoso.com", PrincipalType: 1, IsHiddenInUI: false },
                PrincipalId: 4,
                RoleDefinitionBindings: [{ Id: 1073741829, Name: "Full Control", Description: "", Order: 1, Hidden: false }]
            }
        ];
    }

    public async getLists(excludedLists?: string[]): Promise<IListInfo[]> {
        return [
            { Id: "1", Title: "Campaign Documents", ItemCount: 154, Hidden: false, ItemType: "Library", ServerRelativeUrl: "/sites/marketing/Shared Documents", HasUniqueRoleAssignments: true, EntityTypeName: "Documents" },
            { Id: "2", Title: "Product Specifications", ItemCount: 42, Hidden: false, ItemType: "Library", ServerRelativeUrl: "/sites/marketing/ProductSpecs", HasUniqueRoleAssignments: false, EntityTypeName: "ProductSpecs" },
            { Id: "3", Title: "Event Calendar", ItemCount: 12, Hidden: false, ItemType: "List", ServerRelativeUrl: "/sites/marketing/Lists/Events", HasUniqueRoleAssignments: true, EntityTypeName: "Events" },
            { Id: "4", Title: "Budget Approvals", ItemCount: 8, Hidden: false, ItemType: "List", ServerRelativeUrl: "/sites/marketing/Lists/Budget", HasUniqueRoleAssignments: true, EntityTypeName: "Budget" }
        ];
    }

    public async getListRoleAssignments(listId: string, listTitle: string): Promise<IRoleAssignment[]> {
        return [
            {
                Member: { Id: 2, Title: "Marketing Members", LoginName: "c:0+.w|s-1-5-21...", PrincipalType: 8, IsHiddenInUI: false },
                PrincipalId: 2,
                RoleDefinitionBindings: [{ Id: 1073741830, Name: "Edit", Description: "", Order: 22, Hidden: false }]
            },
            {
                Member: { Id: 5, Title: "External Auditors", LoginName: "c:0-.t|identityprovider|auditors", PrincipalType: 4, IsHiddenInUI: false },
                PrincipalId: 5,
                RoleDefinitionBindings: [{ Id: 1073741826, Name: "Read", Description: "", Order: 2, Hidden: false }]
            }
        ];
    }

    public async getSiteStats(): Promise<ISiteStats> {
        return {
            totalUsers: 142,
            totalGroups: 15,
            uniquePermissionsCount: 8,
            emptyGroupsCount: 2
        };
    }

    public async getSiteGroups(): Promise<IGroup[]> {
        return [
            { Id: 1, Title: "Marketing Owners", LoginName: "owners", Description: "Site Owners", OwnerTitle: "System Account", IsHiddenInUI: false, PrincipalType: 8, UserCount: 1 },
            { Id: 2, Title: "Marketing Members", LoginName: "members", Description: "Site Members", OwnerTitle: "Marketing Owners", IsHiddenInUI: false, PrincipalType: 8, UserCount: 3 },
            { Id: 3, Title: "Marketing Visitors", LoginName: "visitors", Description: "Site Visitors", OwnerTitle: "Marketing Owners", IsHiddenInUI: false, PrincipalType: 8, UserCount: 0 },
            { Id: 10, Title: "Design Team", LoginName: "design", Description: "Designers and Creatives", OwnerTitle: "Marketing Owners", IsHiddenInUI: false, PrincipalType: 8, UserCount: 5 },
            { Id: 11, Title: "Security Audit Team", LoginName: "sec_audit", Description: "External Security Auditors", OwnerTitle: "Active Directory", IsHiddenInUI: false, PrincipalType: 4, UserCount: 0 }
        ];
    }

    public async getGroupMembers(groupId: number): Promise<IUser[]> {
        if (groupId === 1) return [{ Id: 101, Title: "Admin User", Email: "admin@contoso.com", LoginName: "admin", PrincipalType: 1, IsHiddenInUI: false }];
        if (groupId === 2) return [
            { Id: 102, Title: "John Doe", Email: "john.d@contoso.com", LoginName: "johnd", PrincipalType: 1, IsHiddenInUI: false },
            { Id: 103, Title: "Jane Smith", Email: "jane.s@contoso.com", LoginName: "janes", PrincipalType: 1, IsHiddenInUI: false },
            { Id: 104, Title: "Bob Wilson", Email: "bob.w@contoso.com", LoginName: "bobw", PrincipalType: 1, IsHiddenInUI: false }
        ];
        return [];
    }

    public async searchUsers(query: string): Promise<IUser[]> {
        return [
            { Id: 102, Title: "John Doe", Email: "john.d@contoso.com", LoginName: "johnd", PrincipalType: 1, IsHiddenInUI: false },
            { Id: 103, Title: "Jane Smith", Email: "jane.s@contoso.com", LoginName: "janes", PrincipalType: 1, IsHiddenInUI: false }
        ];
    }

    public async getSiteAdmins(): Promise<IUser[]> {
        return [
            { Id: 99, Title: "System Administrator", Email: "admin@contoso.com", LoginName: "admin", PrincipalType: 1, IsHiddenInUI: false, IsSiteAdmin: true }
        ];
    }

    public async getSiteOwners(): Promise<IUser[]> {
        return [
            { Id: 99, Title: "System Administrator", Email: "admin@contoso.com", LoginName: "admin", PrincipalType: 1, IsHiddenInUI: false, IsSiteOwner: true },
            { Id: 100, Title: "Marketing Lead", Email: "lead@contoso.com", LoginName: "lead", PrincipalType: 1, IsHiddenInUI: false, IsSiteOwner: true }
        ];
    }

    public async getCurrentUser(): Promise<IUser> {
        return { Id: 99, Title: "System Administrator", Email: "admin@contoso.com", LoginName: "admin", PrincipalType: 1, IsHiddenInUI: false, IsSiteAdmin: true };
    }

    public async removeSitePermission(principalId: number): Promise<boolean> { return true; }
    public async removeListPermission(listId: string, principalId: number): Promise<boolean> { return true; }
    public async removeItemPermission(listId: string, itemId: number, principalId: number): Promise<boolean> { return true; }
    public async removeUserFromGroup(groupId: number, userId: number): Promise<boolean> {
        console.log(`[Mock] Removing user ${userId} from group ${groupId}`);
        return true;
    }
    public async getUniquePermissionItems(listId: string): Promise<IItemPermission[]> {
        return [
            { Id: 55, Title: "Confidential Strategy.docx", ServerRelativeUrl: "/sites/marketing/docs/strategy.docx", FileSystemObjectType: 1, RoleAssignments: [] }
        ];
    }
    public async getUserGroups(loginName: string): Promise<IGroup[]> {
        // Mock: John Doe is in "Marketing Members" (Id 2)
        if (loginName === "johnd" || loginName.includes("john")) {
            return [
                { Id: 2, Title: "Marketing Members", LoginName: "members", Description: "Site Members", OwnerTitle: "Marketing Owners", IsHiddenInUI: false, PrincipalType: 8, UserCount: 3 }
            ];
        }
        return [];
    }

    public async getExternalUsers(): Promise<IUser[]> {
        return [
            { Id: 201, Title: "Guest User 1", Email: "guest1@partner.com", LoginName: "i:0#.f|membership|guest1@partner.com", PrincipalType: 1, IsHiddenInUI: false },
            { Id: 202, Title: "Vendor Contact", Email: "vendor@contractor.com", LoginName: "i:0#.f|membership|vendor@contractor.com", PrincipalType: 1, IsHiddenInUI: false }
        ];
    }

    public async getSharingLinks(): Promise<ISharingInfo[]> {
        return [
            {
                documentName: "Q1 Financials.xlsx",
                documentUrl: "/sites/marketing/Shared Documents/Q1 Financials.xlsx",
                sharedWith: ["guest1@partner.com", "finance@otherorg.com"],
                linkType: "Specific People"
            },
            {
                documentName: "Brand Guidelines.pdf",
                documentUrl: "/sites/marketing/Shared Documents/Brand Guidelines.pdf",
                sharedWith: ["Everyone in Organization"],
                linkType: "Organization"
            }
        ];
    }

    public async getOrphanedUsers(): Promise<IUser[]> {
        return [
            { Id: 401, Title: "Orphaned User", Email: "", LoginName: String.raw`i:0#.w|domain\orphaned`, PrincipalType: 1, IsHiddenInUI: false }
        ];
    }

    public async getSiteDetails(): Promise<{ Title: string; Url: string; HasUniqueRoleAssignments: boolean }> {
        return {
            Title: "Marketing Site",
            Url: "https://contoso.sharepoint.com/sites/marketing",
            HasUniqueRoleAssignments: true
        };
    }

    public async getSiteUsage(): Promise<ISiteUsage> {
        return {
            storageUsed: 40, // 40%
            storageQuota: 26214400, // 25 TB in MB
            usagePercentage: 0.4,
            lastItemModifiedDate: new Date().toISOString()
        };
    }

    public async getRoleDefinitions(): Promise<IRoleDefinitionDetail[]> {
        return [
            { Id: 1, Name: "Full Control", Description: "Has full control.", BasePermissions: { High: 0, Low: 65535 }, Order: 1, Hidden: false, RoleTypeKind: 5 },
            { Id: 2, Name: "Design", Description: "Can view, add, update, delete, approve, and customize.", BasePermissions: { High: 0, Low: 0 }, Order: 32, Hidden: false, RoleTypeKind: 4 },
            { Id: 3, Name: "Edit", Description: "Can add, edit and delete lists; can view, add, update, and delete list items and documents.", BasePermissions: { High: 0, Low: 0 }, Order: 48, Hidden: false, RoleTypeKind: 6 },
            { Id: 4, Name: "Contribute", Description: "Can view, add, update, and delete list items and documents.", BasePermissions: { High: 0, Low: 0 }, Order: 64, Hidden: false, RoleTypeKind: 3 },
            { Id: 5, Name: "Read", Description: "Can view pages and list items and download documents.", BasePermissions: { High: 0, Low: 1 }, Order: 128, Hidden: false, RoleTypeKind: 2 },
        ];
    }

    public async getUserEffectivePermissions(loginName: string): Promise<string[]> {
        if (loginName.includes("admin")) return ["FullControl", "ManageWeb", "CreateGroups"];
        if (loginName.includes("john")) return ["ViewPages", "EditListItems", "ViewListItems"];
        return ["ViewPages", "ViewListItems"];
    }
}
