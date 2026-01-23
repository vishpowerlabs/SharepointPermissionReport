import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IListInfo, IRoleAssignment } from '../models/IPermissionData';
import { ListCard } from './ListCard';

export interface IListPermissionsProps {
    lists: IListInfo[];
    getListPermissions: (listId: string) => Promise<IRoleAssignment[]>;
    onScanItems: (listId: string) => void;
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
                    onRemovePermission={onRemovePermission ? (pid, pname) => handleRemove(list.Id, pid, pname) : undefined}
                    themeVariant={props.themeVariant}
                    buttonFontSize={props.buttonFontSize}
                    contentFontSize={props.contentFontSize}
                />
            ))}
        </div>
    );
};
