import * as React from 'react';
import { IRoleAssignment, IGroup } from '../models/IPermissionData';
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
    onRemovePermission?: (principalId: number, principalName: string) => void;
    onRemoveFromGroup?: (groupId: number, userId: number, userName: string) => void;
    siteGroups?: IGroup[];
}

const UserGroupCell: React.FunctionComponent<{
    item: any;
    expandedGroups: Set<number>;
    onToggle: (group: IRoleAssignment) => void;
    fontSize?: string;
    onRemovePermission?: (principalId: number, principalName: string) => void;
    onRemoveFromGroup?: (groupId: number, userId: number, userName: string) => void;
    siteGroups?: IGroup[];
}> = ({ item, expandedGroups, onToggle, fontSize, onRemovePermission, onRemoveFromGroup, siteGroups }) => {
    const isGroup = item.Member.PrincipalType === 8; // Only SharePoint Groups are expandable
    const isUser = item.Member.PrincipalType === 1;
    const isExpanded = expandedGroups.has(item.Member.Id);
    const depth = item.depth || 0;

    // Check for empty group
    let isEmptyGroup = false;
    if (isGroup && siteGroups) {
        const groupInfo = siteGroups.find(g => g.Id === item.Member.Id);
        if (groupInfo && groupInfo.UserCount === 0) {
            isEmptyGroup = true;
        }
    }

    const handleDelete = () => {
        if (depth === 0 && onRemovePermission) {
            onRemovePermission(item.Member.Id, item.Member.Title);
        } else if (depth > 0 && onRemoveFromGroup && item.parentGroupId) {
            onRemoveFromGroup(item.parentGroupId, item.Member.Id, item.Member.Title);
        }
    };

    return (
        <div style={{ paddingLeft: `${depth * 24}px`, display: 'flex', alignItems: 'center', fontSize: fontSize, justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
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
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <UserPersona
                            user={item.Member}
                            secondaryText={depth > 0 ? 'Member' : undefined}
                            fontSize={fontSize}
                        />
                        {isEmptyGroup && depth === 0 && (
                            <span style={{
                                background: '#fde7e9',
                                color: '#d13438',
                                fontSize: 10,
                                padding: '2px 6px',
                                borderRadius: 4,
                                fontWeight: 600,
                                border: '1px solid #d13438'
                            }}>Empty</span>
                        )}
                    </div>
                )}
            </div>
            {(
                (isUser && depth === 0 && onRemovePermission) ||
                (depth > 0 && onRemoveFromGroup)
            ) && (
                    <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title={depth === 0 ? "Remove Permission" : "Remove from Group"}
                        onClick={handleDelete}
                        styles={{ root: { height: 24, width: 24, color: '#a80000' } }}
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
    const { permissions, permissionService, onRemovePermission, siteGroups } = props;
    const [expandedGroups, setExpandedGroups] = React.useState<Set<number>>(new Set());
    const [groupMembers, setGroupMembers] = React.useState<{ [key: number]: IRoleAssignment[] }>({});
    const [loadingGroups, setLoadingGroups] = React.useState<Set<number>>(new Set());

    React.useEffect(() => {
        setExpandedGroups(new Set());
        setGroupMembers({});
    }, [permissions]);

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
        const items: (IRoleAssignment & { depth?: number, isLoading?: boolean, parentGroupId?: number })[] = [];

        permissions.forEach(p => {
            items.push({ ...p, depth: 0 });
            const groupId = p.Member.Id;
            if (p.Member.PrincipalType === 8 && expandedGroups.has(groupId)) {
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
                        items.push({ ...m, depth: 1, parentGroupId: groupId });
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
            onRemovePermission={onRemovePermission}
            onRemoveFromGroup={props.onRemoveFromGroup}
            siteGroups={siteGroups}
        />
    ), [expandedGroups, toggleGroup, props.contentFontSize, onRemovePermission, props.onRemoveFromGroup, siteGroups]);

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
