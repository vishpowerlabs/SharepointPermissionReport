import * as React from 'react';
import { IRoleAssignment } from '../models/IPermissionData';
import { UserPersona } from './UserPersona';
import { PermissionBadge } from './PermissionBadge';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IPermissionService } from '../services/IPermissionService';
import styles from './PermissionViewer.module.scss';

export interface ISitePermissionsProps {
    permissions: IRoleAssignment[];
    permissionService?: IPermissionService;
    contentFontSize?: string;
}

const UserGroupCell: React.FunctionComponent<{
    item: any;
    expandedGroups: Set<number>;
    onToggle: (group: IRoleAssignment) => void;
    fontSize?: string;
}> = ({ item, expandedGroups, onToggle, fontSize }) => {
    const isGroup = item.Member.PrincipalType === 8 || item.Member.PrincipalType === 4;
    const isExpanded = expandedGroups.has(item.Member.Id);
    const depth = item.depth || 0;

    return (
        <div style={{ paddingLeft: `${depth * 24}px`, display: 'flex', alignItems: 'center', fontSize: fontSize }}>
            {isGroup && depth === 0 && (
                <IconButton
                    iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                    onClick={() => onToggle(item)}
                    styles={{ root: { height: 24, width: 24, marginRight: 4 } }}
                />
            )}
            {item.isLoading ? (
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontStyle: 'italic', color: '#605e5c' }}>
                    <Spinner size={SpinnerSize.xSmall} />
                    Loading...
                </div>
            ) : (
                <UserPersona
                    user={item.Member}
                    secondaryText={depth > 0 ? 'Member' : undefined}
                    fontSize={fontSize}
                />
            )}
        </div>
    );
};

const PrincipalTypeCell: React.FunctionComponent<{ item: IRoleAssignment; fontSize?: string }> = ({ item, fontSize }) => {
    const typeMap: { [key: number]: string } = { 1: 'User', 4: 'Security Group', 8: 'SharePoint Group' };
    return <span style={{ fontSize: fontSize }}>{typeMap[item.Member.PrincipalType] || 'Unknown'}</span>;
};

const PermissionLevelCell: React.FunctionComponent<{ item: IRoleAssignment; fontSize?: string }> = ({ item, fontSize }) => (
    <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
        {item.RoleDefinitionBindings.map(role => (
            <PermissionBadge key={role.Id} permission={role.Name} fontSize={fontSize ? `calc(${fontSize} - 2px)` : undefined} />
        ))}
    </div>
);

const renderPrincipalType = (item: any, fontSize?: string) => <PrincipalTypeCell item={item} fontSize={fontSize} />;

const renderPermissionLevel = (item: any, fontSize?: string) => <PermissionLevelCell item={item} fontSize={fontSize} />;

export const SitePermissions: React.FunctionComponent<ISitePermissionsProps> = (props) => {
    const { permissions, permissionService } = props;
    const [expandedGroups, setExpandedGroups] = React.useState<Set<number>>(new Set());
    const [groupMembers, setGroupMembers] = React.useState<{ [key: number]: IRoleAssignment[] }>({});
    const [loadingGroups, setLoadingGroups] = React.useState<Set<number>>(new Set());

    const toggleGroup = async (group: IRoleAssignment) => {
        const groupId = group.Member.Id;
        const newExpanded = new Set(expandedGroups);

        if (newExpanded.has(groupId)) {
            newExpanded.delete(groupId);
            setExpandedGroups(newExpanded);
            return;
        }

        newExpanded.add(groupId);
        setExpandedGroups(newExpanded);

        if (!groupMembers[groupId] && permissionService) {
            setLoadingGroups(prev => new Set(prev).add(groupId));
            try {
                const users = await permissionService.getGroupMembers(groupId);
                // Convert users to IRoleAssignment structure for display
                const memberAssignments: IRoleAssignment[] = users.map(u => ({
                    Member: u,
                    PrincipalId: u.Id,
                    RoleDefinitionBindings: group.RoleDefinitionBindings // Inherit permission levels from group
                }));
                setGroupMembers(prev => ({ ...prev, [groupId]: memberAssignments }));
            } catch (err) {
                console.error(err);
            } finally {
                setLoadingGroups(prev => {
                    const next = new Set(prev);
                    next.delete(groupId);
                    return next;
                });
            }
        }
    };

    const displayItems = React.useMemo(() => {
        const items: (IRoleAssignment & { depth?: number, isLoading?: boolean })[] = [];

        permissions.forEach(p => {
            items.push({ ...p, depth: 0 });
            const groupId = p.Member.Id;
            if ((p.Member.PrincipalType === 8 || p.Member.PrincipalType === 4) && expandedGroups.has(groupId)) {
                if (loadingGroups.has(groupId)) {
                    // Placeholder for loading
                    items.push({
                        Member: { Id: -1, Title: 'Loading...', IsHiddenInUI: false, PrincipalType: 1 },
                        PrincipalId: -1,
                        RoleDefinitionBindings: [],
                        depth: 1,
                        isLoading: true
                    } as any);
                } else if (groupMembers[groupId]) {
                    groupMembers[groupId].forEach(m => {
                        items.push({ ...m, depth: 1 });
                    });
                }
            }
        });
        return items;
    }, [permissions, expandedGroups, groupMembers, loadingGroups]);

    const renderUserGroup = React.useCallback((item: any) => (
        <UserGroupCell
            item={item}
            expandedGroups={expandedGroups}
            onToggle={toggleGroup}
            fontSize={props.contentFontSize}
        />
    ), [expandedGroups, toggleGroup, props.contentFontSize]);

    const columns: IColumn[] = React.useMemo(() => [
        {
            key: 'user',
            name: 'User/Group',
            fieldName: 'Member',
            minWidth: 200,
            maxWidth: 400,
            onRender: renderUserGroup
        },
        {
            key: 'type',
            name: 'Type',
            fieldName: 'Member',
            minWidth: 100,
            maxWidth: 150,
            onRender: (item) => renderPrincipalType(item, props.contentFontSize)
        },
        {
            key: 'level',
            name: 'Permission Level',
            fieldName: 'RoleDefinitionBindings',
            minWidth: 150,
            maxWidth: 200,
            onRender: (item) => renderPermissionLevel(item, props.contentFontSize)
        }
    ], [renderUserGroup, props.contentFontSize]);

    return (
        <div className={styles.content}>
            <div className={styles.permissionTable} style={{ border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden' }}>
                <DetailsList
                    items={displayItems}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    styles={{
                        root: { background: '#ffffff', fontSize: props.contentFontSize },
                        headerWrapper: { background: '#faf9f8' },
                        contentWrapper: { fontSize: props.contentFontSize }
                    }}
                />
            </div>
        </div>
    );
};
