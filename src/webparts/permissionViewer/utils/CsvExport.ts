import { IRoleAssignment, IListInfo, IItemPermission, IPublicAccessResult, IOversharedFolder } from '../models/IPermissionData';
import { IPermissionService } from '../services/IPermissionService';

const CSV_HEADER_URI = "data:text/csv;charset=utf-8,";

export const downloadCsv = (csvContent: string, fileName: string): void => {
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", fileName);
    document.body.appendChild(link);
    link.click();
    link.remove();
};

const getPrincipalTypeLabel = (pType: number): string => {
    if (pType === 1) return 'User';
    if (pType === 8) return 'SharePoint Group';
    return 'Security Group';
};

const getMemberTypeLabel = (pType: number): string => {
    if (pType === 1) return 'User';
    if (pType === 4) return 'Security Group';
    return 'Group';
};

export const exportSitePermissions = async (
    permissions: IRoleAssignment[],
    permissionService: IPermissionService
): Promise<void> => {
    let csvContent = CSV_HEADER_URI;
    csvContent += "Type,Name,Email,Permission Level,Member Of\n";

    for (const p of permissions) {
        const type = getPrincipalTypeLabel(p.Member.PrincipalType);
        const roles = p.RoleDefinitionBindings.map(r => r.Name).join('; ');
        csvContent += `${type},"${p.Member.Title}","${p.Member.Email || ''}","${roles}",\n`;

        if (p.Member.PrincipalType === 8) {
            try {
                const members = await permissionService.getGroupMembers(p.Member.Id);
                members.forEach(m => {
                    const memberType = getMemberTypeLabel(m.PrincipalType);
                    csvContent += `${memberType},"${m.Title}","${m.Email || ''}","${roles}","${p.Member.Title}"\n`;
                });
            } catch (e) {
                console.error(`Failed to fetch members for group ${p.Member.Title}`, e);
            }
        }
    }
    downloadCsv(csvContent, 'SitePermissions.csv');
};

export const exportListPermissions = async (
    lists: IListInfo[],
    permissionService: IPermissionService
): Promise<void> => {
    let csvContent = CSV_HEADER_URI;
    csvContent += "List Name,Url,Type,Has Unique Permissions,User/Group,User Type,Permission Level\n";

    for (const list of lists) {
        let permissions: IRoleAssignment[] = [];
        if (list.HasUniqueRoleAssignments) {
            try {
                permissions = await permissionService.getListRoleAssignments(list.Id, list.Title);
            } catch (e) {
                console.error(`Failed to export permissions for ${list.Title}`, e);
            }
        }

        const listType = list.ItemType;
        const cleanTitle = list.Title.replaceAll('"', '""');

        if (!list.HasUniqueRoleAssignments) {
            csvContent += `"${cleanTitle}","${list.ServerRelativeUrl}",${listType},No,Inherited,Inherited,Inherited\n`;
        } else if (permissions.length === 0) {
            csvContent += `"${cleanTitle}","${list.ServerRelativeUrl}",${listType},Yes,No Permissions Found,-,-\n`;
        } else {
            permissions.forEach(p => {
                const userType = getPrincipalTypeLabel(p.Member.PrincipalType);
                const roles = p.RoleDefinitionBindings.map(r => r.Name).join('; ');
                const memberName = p.Member.Title.replaceAll('"', '""');
                csvContent += `"${cleanTitle}","${list.ServerRelativeUrl}",${listType},Yes,"${memberName}",${userType},"${roles}"\n`;
            });
        }
    }
    downloadCsv(csvContent, 'ListPermissions.csv');
};

export const exportDeepScanResults = (
    items: IItemPermission[],
    listTitle: string
): void => {
    let csvContent = CSV_HEADER_URI;
    csvContent += "Path,Name,Type,User/Group,Permission Level\n";

    items.forEach(item => {
        const typeStr = item.FileSystemObjectType === 1 ? 'Folder' : 'File';
        const cleanPath = item.ServerRelativeUrl;
        const cleanName = item.Title.replaceAll('"', '""');

        if (item.RoleAssignments.length === 0) {
            csvContent += `"${cleanPath}","${cleanName}",${typeStr},No Permissions Found,-\n`;
        } else {
            item.RoleAssignments.forEach(p => {
                const roles = p.RoleDefinitionBindings.map(r => r.Name).join('; ');
                const memberName = p.Member.Title.replaceAll('"', '""');
                csvContent += `"${cleanPath}","${cleanName}",${typeStr},"${memberName}","${roles}"\n`;
            });
        }
    });

    downloadCsv(csvContent, `${listTitle}_DeepScan.csv`);
};

export const exportStorageMetrics = (
    lists: IListInfo[],
    siteUsage: any // using any or ISiteUsage if imported
): void => {
    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += "List Name,Item Count,Total Size (Bytes),% of Site,Last Modified\n";

    const totalUsed = siteUsage?.storageUsed || 1; // Avoid divide by zero

    lists.forEach(list => {
        const size = list.TotalSize || 0;
        const percent = (size / totalUsed) * 100;
        const cleanTitle = list.Title.replaceAll('"', '""');
        const lastMod = list.LastItemModifiedDate ? new Date(list.LastItemModifiedDate).toLocaleDateString() : '-';

        csvContent += `"${cleanTitle}",${list.ItemCount},${size},${percent.toFixed(2)}%,${lastMod}\n`;
    });

    downloadCsv(csvContent, 'StorageMetrics.csv');
};

export const exportPublicAccessResults = (
    items: IPublicAccessResult[]
): void => {
    let csvContent = CSV_HEADER_URI;
    csvContent += "Scope,Location,Url,Public Group,Type,Permission\n";

    items.forEach(item => {
        const cleanName = item.itemName.replaceAll('"', '""');
        const roles = item.roles.join('; ');
        csvContent += `${item.scope},"${cleanName}","${item.itemUrl}","${item.principalName}",${item.principalType},"${roles}"\n`;
    });

    downloadCsv(csvContent, 'PublicAccessReport.csv');
};

export const exportFolderOversharingResults = (
    items: IOversharedFolder[],
    listTitle: string
): void => {
    let csvContent = CSV_HEADER_URI;
    csvContent += "Folder Name,Shared With,Permissions,Path,Login Name\n";

    items.forEach(item => {
        const cleanName = item.Name.replaceAll('"', '""');
        const cleanPath = item.Path;
        const sharedWith = item.SharedWith.replaceAll('"', '""');
        const perms = item.Permissions.replaceAll('"', '""');
        const login = (item.LoginName || '').replaceAll('"', '""');

        csvContent += `"${cleanName}","${sharedWith}","${perms}","${cleanPath}","${login}"\n`;
    });

    downloadCsv(csvContent, `${listTitle}_OversharedFolders.csv`);
};
