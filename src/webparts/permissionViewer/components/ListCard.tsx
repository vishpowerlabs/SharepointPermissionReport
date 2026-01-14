import * as React from 'react';
import { IListInfo, IRoleAssignment } from '../models/IPermissionData';
import { PermissionBadge } from './PermissionBadge';
import styles from './PermissionViewer.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ListPermissionsTable } from './ListPermissionsTable';

export interface IListCardProps {
    list: IListInfo;
    permissions: IRoleAssignment[];
    isLoading: boolean;
    isExpanded: boolean;
    onExpand: () => void;
    onScanItems?: (listId: string) => void;
    themeVariant: IReadonlyTheme | undefined;
}

export const ListCard: React.FunctionComponent<IListCardProps> = (props) => {
    const { list, permissions, isLoading, isExpanded, onExpand, themeVariant } = props;

    // Theme colors
    const primaryColor = themeVariant?.palette?.themePrimary || '#0078d4';
    const textColor = themeVariant?.palette?.neutralPrimary || '#323130';
    const borderColor = themeVariant?.palette?.neutralLight || '#e1dfdd';


    // Render permissions table rows using plain HTML/CSS as per mockup structure for simple list cards, 
    // or use DetailsList. The mockup shows a table inside the card.

    const renderPermissionsContent = () => {
        if (!list.HasUniqueRoleAssignments) {
            return (
                <div className={styles.inheritedInfo} style={{ background: '#e6f2ff', borderLeft: '4px solid #0078d4', borderRadius: '4px', padding: '16px', display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <span style={{ color: '#0078d4', fontSize: '20px' }}>ℹ️</span>
                    <div>
                        <div style={{ fontSize: '14px', color: '#323130' }}>This list inherits permissions from the parent site</div>
                        <button
                            onClick={() => {/* Navigate to site perms */ }}
                            style={{
                                background: 'none',
                                border: 'none',
                                padding: 0,
                                color: '#0078d4',
                                fontSize: '12px',
                                textDecoration: 'underline',
                                cursor: 'pointer'
                            }}>
                            View Site Permissions
                        </button>
                    </div>
                </div>
            );
        }

        if (isLoading) {
            return <div>Loading permissions...</div>;
        }

        return <ListPermissionsTable permissions={permissions} />;
    };

    return (
        <div className={styles.listCard} style={{ background: '#ffffff', border: `1px solid ${borderColor}`, borderRadius: '8px', padding: '20px', marginBottom: '16px' }}>
            <div className={styles.listCardHeader} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                <button className={styles.listTitle}
                    disabled={!list.HasUniqueRoleAssignments}
                    style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '12px',
                        cursor: list.HasUniqueRoleAssignments ? 'pointer' : 'default',
                        color: textColor,
                        background: 'transparent',
                        border: 'none',
                        padding: 0,
                        textAlign: 'left'
                    }}
                    onClick={list.HasUniqueRoleAssignments ? onExpand : undefined}>
                    <span className={styles.listIcon} style={{ fontSize: '20px', color: primaryColor }}>
                        {list.ItemType === 'Library' ? '📁' : '📋'}
                    </span>
                    <h3 style={{ fontSize: '18px', fontWeight: 600, color: textColor, margin: 0 }}>{list.Title}</h3>
                    {list.HasUniqueRoleAssignments && (
                        <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} />
                    )}
                </button>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {props.onScanItems && (
                        <button
                            onClick={(e) => { e.stopPropagation(); props.onScanItems!(list.Id); }}
                            style={{
                                background: 'transparent',
                                border: `1px solid ${primaryColor}`,
                                color: primaryColor,
                                padding: '4px 12px',
                                borderRadius: '4px',
                                cursor: 'pointer',
                                fontSize: '12px',
                                fontWeight: 600,
                                width: '140px',
                                minWidth: '140px',
                                display: 'flex',
                                justifyContent: 'center',
                                alignItems: 'center'
                            }}
                            title="Scan all files/folders in in this list for specific unique permissions">
                            🔍 Deep Scan
                        </button>
                    )}
                    <PermissionBadge permission={list.HasUniqueRoleAssignments ? 'Unique' : 'Inherited'} isInheritanceStatus={true} />
                </div>
            </div>
            <div className={styles.listUrl} style={{ fontSize: '12px', color: '#605e5c', fontStyle: 'italic', marginBottom: '16px' }}>
                {list.ServerRelativeUrl}
            </div>

            {isExpanded && (
                <div className="list-permissions-content">
                    {renderPermissionsContent()}
                </div>
            )}
        </div>
    );
};
