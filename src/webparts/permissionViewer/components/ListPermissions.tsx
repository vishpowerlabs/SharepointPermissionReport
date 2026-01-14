import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IListInfo, IRoleAssignment } from '../models/IPermissionData';
import { ListCard } from './ListCard';

export interface IListPermissionsProps {
    lists: IListInfo[];
    getListPermissions: (listId: string) => Promise<IRoleAssignment[]>;
    onScanItems: (listId: string) => void;
    themeVariant: IReadonlyTheme | undefined;

    buttonFontSize?: string;
    contentFontSize?: string;
}

export const ListPermissions: React.FunctionComponent<IListPermissionsProps> = (props) => {
    const { lists, getListPermissions } = props;
    const [expandedList, setExpandedList] = React.useState<string | null>(null);
    const [permissionsMap, setPermissionsMap] = React.useState<{ [key: string]: IRoleAssignment[] }>({});
    const [loadingMap, setLoadingMap] = React.useState<{ [key: string]: boolean }>({});

    const handleExpand = async (listId: string) => {
        const list = lists.find(l => l.Id === listId);
        if (!list?.HasUniqueRoleAssignments) return;

        if (expandedList === listId) {
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
                    themeVariant={props.themeVariant}
                    buttonFontSize={props.buttonFontSize}
                    contentFontSize={props.contentFontSize}
                />
            ))}
        </div>
    );
};
