import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { IUser, IRoleAssignment } from '../models/IPermissionData';
import { UserPersona } from './UserPersona';
import styles from './PermissionViewer.module.scss';

export interface IOrphanedUsersPanelProps {
    isOpen: boolean;
    onDismiss: () => void;
    orphanedUsers: IRoleAssignment[]; // Passing RoleAssignments to have context of permissions if needed, or at least the User object
    onRemoveUser: (user: IUser) => Promise<void>;
    isRemoving: boolean;
    contentFontSize?: string;
}

export const OrphanedUsersPanel: React.FunctionComponent<IOrphanedUsersPanelProps> = (props) => {
    const { isOpen, onDismiss, orphanedUsers, onRemoveUser, isRemoving, contentFontSize } = props;
    const [removingUserId, setRemovingUserId] = React.useState<number | null>(null);

    const handleRemove = async (user: IUser) => {
        setRemovingUserId(user.Id);
        await onRemoveUser(user);
        setRemovingUserId(null);
    };

    const columns: IColumn[] = [
        {
            key: 'user',
            name: 'User',
            fieldName: 'Member',
            minWidth: 150,
            maxWidth: 200,
            onRender: (item: IRoleAssignment) => (
                <UserPersona user={item.Member} fontSize={contentFontSize} />
            )
        },

        {
            key: 'status',
            name: 'Reason',
            fieldName: 'OrphanStatus',
            minWidth: 100,
            maxWidth: 150,
            onRender: (item: IRoleAssignment) => {
                const status = item.Member.OrphanStatus;
                if (status === 'Deleted') {
                    return <span style={{ color: '#a80000', fontWeight: 600, fontSize: contentFontSize }}>Deleted from AD</span>;
                } else if (status === 'Disabled') {
                    return <span style={{ color: '#a80000', fontWeight: 600, fontSize: contentFontSize }}>Account Disabled</span>;
                }
                return <span style={{ fontSize: contentFontSize }}>Unknown</span>;
            }
        },
        {
            key: 'actions',
            name: 'Actions',
            minWidth: 100,
            maxWidth: 100,
            onRender: (item: IRoleAssignment) => (
                <PrimaryButton
                    text={removingUserId === item.Member.Id ? "Removing..." : "Remove"}
                    onClick={() => handleRemove(item.Member)}
                    disabled={isRemoving || removingUserId !== null}
                    styles={{
                        root: { height: 24, fontSize: contentFontSize, padding: '0 8px', backgroundColor: '#d13438', border: 'none' },
                        rootHovered: { backgroundColor: '#a80000' },
                        label: { fontSize: '11px', fontWeight: 600 }
                    }}
                />
            )
        }
    ];

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.custom}
            customWidth="600px"
            headerText="Orphaned Users Detected"
            closeButtonAriaLabel="Close"
            isBlocking={false}
        >
            <div style={{ marginTop: '20px' }}>
                <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 20 } }}>
                    The following users were found to be <strong>Deleted</strong> or <strong>Disabled</strong> in Azure AD but still have permissions on this site.
                    Review and remove them if necessary.
                </MessageBar>

                {orphanedUsers.length === 0 ? (
                    <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>
                        No orphaned users found in the current selection.
                    </div>
                ) : (
                    <DetailsList
                        items={orphanedUsers}
                        columns={columns}
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.justified}
                        styles={{ root: { marginTop: 10 } }}
                    />
                )}
            </div>
        </Panel>
    );
};
