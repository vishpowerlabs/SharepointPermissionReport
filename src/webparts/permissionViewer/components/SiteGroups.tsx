import * as React from 'react';
import { DetailsList, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Text } from '@fluentui/react/lib/Text';
import { IconButton, PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { IGroup, IUser } from '../models/IPermissionData';
import { IPermissionService } from '../services/IPermissionService';

export interface ISiteGroupsProps {
    groups: IGroup[];
    isLoading: boolean;
    permissionService: IPermissionService;
    contentFontSize?: string;
}

export const SiteGroups: React.FunctionComponent<ISiteGroupsProps> = (props) => {
    const [filteredGroups, setFilteredGroups] = React.useState<IGroup[]>([]);

    // Panel State
    const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
    const [selectedGroup, setSelectedGroup] = React.useState<IGroup | null>(null);
    const [groupMembers, setGroupMembers] = React.useState<IUser[]>([]);
    const [isLoadingMembers, setIsLoadingMembers] = React.useState<boolean>(false);

    // Remote User Dialog State
    const [isDeleteDialogHidden, setIsDeleteDialogHidden] = React.useState<boolean>(true);
    const [userToDelete, setUserToDelete] = React.useState<IUser | null>(null);
    const [isDeleting, setIsDeleting] = React.useState<boolean>(false);

    React.useEffect(() => {
        setFilteredGroups(props.groups);
    }, [props.groups]);

    const onSearch = (newValue: string) => {
        if (!newValue) {
            setFilteredGroups(props.groups);
            return;
        }

        const lowerQuery = newValue.toLowerCase();
        const filtered = props.groups.filter(g =>
            g.Title.toLowerCase().includes(lowerQuery) ||
            (g.Description && g.Description.toLowerCase().includes(lowerQuery))
        );
        setFilteredGroups(filtered);
    };

    const onGroupClick = (group: IGroup) => {
        setSelectedGroup(group);
        setIsPanelOpen(true);
        loadGroupMembers(group.Id);
    };

    const loadGroupMembers = async (groupId: number) => {
        setIsLoadingMembers(true);
        setGroupMembers([]);
        try {
            const members = await props.permissionService.getGroupMembers(groupId);
            setGroupMembers(members);
        } catch (error) {
            console.error("Error loading members", error);
        } finally {
            setIsLoadingMembers(false);
        }
    };

    const onDeleteUserClick = (user: IUser) => {
        setUserToDelete(user);
        setIsDeleteDialogHidden(false);
    };

    const confirmDeleteUser = async () => {
        if (!selectedGroup || !userToDelete) return;

        setIsDeleting(true);
        try {
            const success = await props.permissionService.removeUserFromGroup(selectedGroup.Id, userToDelete.Id);
            if (success) {
                // Remove from local state to avoid full reload
                setGroupMembers(prev => prev.filter(u => u.Id !== userToDelete.Id));
                setIsDeleteDialogHidden(true);
                setUserToDelete(null);
            } else {
                alert("Failed to remove user. Please try again.");
            }
        } catch (error) {
            console.error("Error removing user", error);
            alert("Error removing user.");
        } finally {
            setIsDeleting(false);
        }
    };

    const columns: IColumn[] = [
        {
            key: 'Title',
            name: 'Group Name',
            fieldName: 'Title',
            minWidth: 150,
            maxWidth: 250,
            isResizable: true,
            onRender: (item: IGroup) => (
                <span
                    style={{ fontSize: props.contentFontSize, fontWeight: 600, color: '#0078d4', cursor: 'pointer' }}
                    onClick={() => onGroupClick(item)}
                >
                    {item.Title}
                </span>
            )
        },
        {
            key: 'Description',
            name: 'Description',
            fieldName: 'Description',
            minWidth: 200,
            maxWidth: 400,
            isResizable: true,
            onRender: (item: IGroup) => <span style={{ fontSize: props.contentFontSize }}>{item.Description}</span>
        },
        {
            key: 'Owner',
            name: 'Group Owner',
            fieldName: 'OwnerTitle', // Map to OwnerTitle if available, logic might need check in Service
            minWidth: 100,
            isResizable: true,
            onRender: (item: IGroup) => <span style={{ fontSize: props.contentFontSize }}>{item.OwnerTitle}</span>
        }
    ];

    return (
        <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }} style={{ marginBottom: 20 }}>
                <SearchBox
                    placeholder="Search groups..."
                    styles={{ root: { width: 300 } }}
                    onChange={(_, val) => onSearch(val || '')}
                />
            </Stack>

            {props.isLoading ? (
                <Spinner size={SpinnerSize.large} label="Loading site groups..." />
            ) : (
                <>
                    {filteredGroups.length === 0 ? (
                        <MessageBar messageBarType={MessageBarType.info}>No groups found.</MessageBar>
                    ) : (
                        <DetailsList
                            items={filteredGroups}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            onItemInvoked={onGroupClick} // Handle double click or Enter
                            styles={{ root: { selectors: { '.ms-DetailsRow': { cursor: 'pointer' } } } }}
                        />
                    )}
                </>
            )}

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                type={PanelType.medium}
                headerText={selectedGroup ? selectedGroup.Title : "Group Members"}
                closeButtonAriaLabel="Close"
            >
                {isLoadingMembers ? (
                    <Spinner size={SpinnerSize.medium} label="Loading members..." style={{ marginTop: 20 }} />
                ) : (
                    <Stack tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                        {groupMembers.length === 0 ? (
                            <Text>No members in this group.</Text>
                        ) : (
                            groupMembers.map(member => (
                                <div key={member.Id || member.LoginName} style={{ display: 'flex', alignItems: 'center', padding: '10px', borderBottom: '1px solid #eee' }}>
                                    <Persona
                                        text={member.Title}
                                        secondaryText={member.Email}
                                        size={PersonaSize.size40}
                                        showSecondaryText={true}
                                    />
                                    <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <span style={{ fontSize: '12px', color: '#666' }}>
                                            {member.PrincipalType === 1 ? 'User' : 'Group'}
                                        </span>
                                        <IconButton
                                            iconProps={{ iconName: 'Delete' }}
                                            title="Remove User"
                                            ariaLabel="Remove User"
                                            disabled={isDeleting}
                                            onClick={() => onDeleteUserClick(member)}
                                            styles={{ root: { color: '#a80000', '&:hover': { color: '#d80000' } } }}
                                        />
                                    </div>
                                </div>
                            ))
                        )}
                    </Stack>
                )}
            </Panel>

            <Dialog
                hidden={isDeleteDialogHidden}
                onDismiss={() => { if (!isDeleting) setIsDeleteDialogHidden(true); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Remove User',
                    subText: `Are you sure you want to remove '${userToDelete?.Title}' from group '${selectedGroup?.Title}'?`
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={confirmDeleteUser} text="Remove" disabled={isDeleting} />
                    <DefaultButton onClick={() => setIsDeleteDialogHidden(true)} text="Cancel" disabled={isDeleting} />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
