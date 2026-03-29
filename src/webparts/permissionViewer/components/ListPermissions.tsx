import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IListInfo, IRoleAssignment } from '../models/IPermissionData';
import { ListCard } from './ListCard';

export interface IListPermissionsProps {
    lists: IListInfo[];
    getListPermissions: (listId: string) => Promise<IRoleAssignment[]>;
    onScanItems: (listId: string) => void;
    onCheckOrphans?: (listId: string, currentPerms: IRoleAssignment[]) => Promise<IRoleAssignment[]>;
    onCheckFolders?: (listId: string) => void;
    onCheckRootItems?: (listId: string) => void;

    themeVariant: IReadonlyTheme | undefined;
    onRemovePermission?: (listId: string, principalId: number, principalName: string) => Promise<boolean>;

    buttonFontSize?: string;
    contentFontSize?: string;
    forcedExpandedListId?: string | null;
}

export const ListPermissions: React.FunctionComponent<IListPermissionsProps> = (props) => {
    const { lists, getListPermissions, onRemovePermission, forcedExpandedListId } = props;
    const [expandedList, setExpandedList] = React.useState<string | null>(null);
    const [permissionsMap, setPermissionsMap] = React.useState<{ [key: string]: IRoleAssignment[] }>({});
    const [loadingMap, setLoadingMap] = React.useState<{ [key: string]: boolean }>({});

    // Effect to handle forced expansion from parent (Demo Mode)
    React.useEffect(() => {
        if (forcedExpandedListId) {
            handleExpand(forcedExpandedListId);
        }
    }, [forcedExpandedListId]);

    const handleExpand = async (listId: string) => {
        const list = lists.find(l => l.Id === listId);
        if (!list?.HasUniqueRoleAssignments) return;

        // Toggle logic: if clicking already expanded, close it. BUT if forced, we generally want to open.
        // For simplicity reusing logic but handling the check.
        if (expandedList === listId && !forcedExpandedListId) {
            setExpandedList(null);
            return;
        }

        setExpandedList(listId);

        if (!permissionsMap[listId]) {
            setLoadingMap(prev => ({ ...prev, [listId]: true }));
            const perms = await getListPermissions(listId);
            setPermissionsMap(prev => ({ ...prev, [listId]: perms }));
            setLoadingMap(prev => ({ ...prev, [listId]: false }));
        }
    };

    const handleRemove = async (listId: string, principalId: number, principalName: string) => {
        if (!onRemovePermission) return;
        const success = await onRemovePermission(listId, principalId, principalName);
        if (success) {
            // Refresh permissions for this list
            setLoadingMap(prev => ({ ...prev, [listId]: true }));
            const perms = await getListPermissions(listId);
            setPermissionsMap(prev => ({ ...prev, [listId]: perms }));
            setLoadingMap(prev => ({ ...prev, [listId]: false }));
        }
    };

    const handleCheckOrphans = async (listId: string) => {
        // Can't easily access permissionService here directly as it's not a prop, but getListPermissions is.
        // Needs a new prop or a way to call service. 
        // Actually, ListPermissions DOES NOT receive permissionService. 
        // It receives getListPermissions. 
        // I need to add onCheckOrphans prop to ListPermissions and pass it down.
        // Wait, I added onCheckOrphans to ListCardProps but ListPermissions needs to implement it.
        // But ListPermissions delegates to its parent (PermissionViewer) usually for data fetching via props?
        // No, `getListPermissions` is passed from parent.
        // I should add `checkOrphans: (listId: string, currentPerms: IRoleAssignment[]) => Promise<IRoleAssignment[]>` or similar to props?
        // Or just `onCheckOrphans`.

        if (props.onCheckOrphans) {
            setLoadingMap(prev => ({ ...prev, [listId]: true }));
            try {
                let currentPerms = permissionsMap[listId];

                // If permissions not loaded (collapsed state), fetch them first
                if (!currentPerms || currentPerms.length === 0) {
                    currentPerms = await getListPermissions(listId);
                    // Update map so we have them shown if expanded later
                    setPermissionsMap(prev => ({ ...prev, [listId]: currentPerms }));
                }

                if (currentPerms && currentPerms.length > 0) {
                    const updatedPerms = await props.onCheckOrphans(listId, currentPerms);
                    // Force clone for Member updates
                    const finalPerms = updatedPerms.map(p => ({ ...p, Member: { ...p.Member } }));
                    setPermissionsMap(prev => ({ ...prev, [listId]: finalPerms }));

                    // Also expand the list to show results if not expanded? 
                    // Maybe better to just let user expand, or auto-expand.
                    // Let's auto-expand so they see the result.
                    if (expandedList !== listId) {
                        setExpandedList(listId);
                    }
                }
            } catch (error) {
                console.error("Error checking orphans for list " + listId, error);
            } finally {
                setLoadingMap(prev => ({ ...prev, [listId]: false }));
            }
        }
    };

    return (
        <div>
            {lists.map(list => (
                <ListCard
                    key={list.Id}
                    list={list}
                    permissions={permissionsMap[list.Id] || []}
                    isLoading={loadingMap[list.Id] || false}
                    isExpanded={expandedList === list.Id}
                    onExpand={() => handleExpand(list.Id)}
                    onScanItems={() => props.onScanItems(list.Id)}
                    onCheckOrphans={props.onCheckOrphans ? () => handleCheckOrphans(list.Id) : undefined}
                    onCheckFolders={props.onCheckFolders ? () => props.onCheckFolders!(list.Id) : undefined}
                    onCheckRootItems={props.onCheckRootItems ? () => props.onCheckRootItems!(list.Id) : undefined}

                    onRemovePermission={onRemovePermission ? (pid, pname) => handleRemove(list.Id, pid, pname) : undefined}
                    themeVariant={props.themeVariant}
                    buttonFontSize={props.buttonFontSize}
                    contentFontSize={props.contentFontSize}
                />
            ))}
        </div>
    );
};
