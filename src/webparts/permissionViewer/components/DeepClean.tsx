import * as React from 'react';
import { PrimaryButton, DefaultButton, ProgressIndicator, MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, Stack } from '@fluentui/react';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode, IGroup } from '@fluentui/react/lib/DetailsList';
import { IPermissionService } from '../services/IPermissionService';
import { IUser, IListInfo } from '../models/IPermissionData';
import { UserPersona } from './UserPersona';
import styles from './PermissionViewer.module.scss';

export interface IDeepCleanProps {
    permissionService: IPermissionService;
    lists: IListInfo[];
    contentFontSize?: string;
}

interface IOrphanResult {
    key: string;
    listName: string;
    listId: string;
    itemId?: number; // If item level
    itemName: string;
    itemUrl: string;
    scope: 'List' | 'Item' | 'Folder';
    user: IUser;
    roles: string[];
    status: string; // 'Deleted' | 'Disabled'
}

export const DeepClean: React.FunctionComponent<IDeepCleanProps> = (props) => {
    const { permissionService, lists, contentFontSize } = props;

    const [isScanning, setIsScanning] = React.useState(false);
    const [progress, setProgress] = React.useState<number>(0);
    const [progressLabel, setProgressLabel] = React.useState<string>('');
    const [orphans, setOrphans] = React.useState<IOrphanResult[]>([]);
    const [groups, setGroups] = React.useState<IGroup[]>([]);

    // Removal State
    const [itemToRemove, setItemToRemove] = React.useState<IOrphanResult | null>(null);
    const [isRemoving, setIsRemoving] = React.useState(false);
    const [hideDialog, setHideDialog] = React.useState(true);
    const [removalMessage, setRemovalMessage] = React.useState<string | null>(null);

    // Abort Controller for stopping scan
    const abortController = React.useRef<AbortController | null>(null);

    const stopScan = () => {
        if (abortController.current) {
            abortController.current.abort();
        }
    };

    const runDeepClean = async () => {
        setIsScanning(true);
        setOrphans([]);
        setGroups([]);
        setProgress(0);
        setRemovalMessage(null);

        // Init AbortController
        abortController.current = new AbortController();
        const signal = abortController.current.signal;

        const foundOrphans: IOrphanResult[] = [];
        const uniqueLists = lists.filter(l => l.HasUniqueRoleAssignments).sort((a, b) => a.Title.localeCompare(b.Title));
        const totalSteps = uniqueLists.length;

        try {
            for (let i = 0; i < uniqueLists.length; i++) {
                if (signal.aborted) throw new Error('Scan aborted by user.');

                const list = uniqueLists[i];
                setProgress((i / totalSteps));
                setProgressLabel(`Scanning ${list.Title}...`);

                try {
                    // 1. Check List Level Orphans
                    const listPerms = await permissionService.getListRoleAssignments(list.Id, list.Title);

                    if (signal.aborted) throw new Error('Scan aborted by user.');

                    // identify unique users to check
                    const listUsersToCheck = listPerms
                        .filter(p => p.Member.PrincipalType === 1) // Users only
                        .map(p => p.Member as IUser);

                    if (listUsersToCheck.length > 0) {
                        const checkedUsers = await permissionService.checkOrphanUsers(listUsersToCheck);

                        listPerms.forEach(p => {
                            const statusUser = checkedUsers.find(u => u.Id === p.Member.Id);
                            if (statusUser && statusUser.OrphanStatus === 'Deleted') {
                                foundOrphans.push({
                                    key: `list-${list.Id}-${p.Member.Id}`,
                                    listName: list.Title,
                                    listId: list.Id,
                                    itemName: list.Title,
                                    itemUrl: list.ServerRelativeUrl,
                                    scope: 'List',
                                    user: { ...p.Member, OrphanStatus: statusUser.OrphanStatus } as IUser,
                                    roles: p.RoleDefinitionBindings.map(r => r.Name),
                                    status: statusUser.OrphanStatus
                                });
                            }
                        });
                    }

                    if (signal.aborted) throw new Error('Scan aborted by user.');

                    // 2. Check Item Level Orphans
                    // Pass signal to service
                    const items = await permissionService.getUniquePermissionItems(list.Id, signal);

                    // Collect all users from all items to batch check
                    const allItemUsersMap = new Map<string, IUser>();
                    items.forEach(item => {
                        item.RoleAssignments.forEach(ra => {
                            if (ra.Member.PrincipalType === 1) {
                                // Fix Map Key type
                                const key = ra.Member.LoginName || ra.Member.Email || ra.Member.Title;
                                allItemUsersMap.set(key, ra.Member as IUser);
                            }
                        });
                    });

                    if (allItemUsersMap.size > 0) {
                        const itemDestUsers = Array.from(allItemUsersMap.values());
                        const checkedItemUsers = await permissionService.checkOrphanUsers(itemDestUsers);
                        const checkedUserMap = new Map(checkedItemUsers.map(u => [u.LoginName || u.Email || u.Title, u]));

                        items.forEach(item => {
                            item.RoleAssignments.forEach(ra => {
                                if (ra.Member.PrincipalType === 1) {
                                    const key = ra.Member.LoginName || ra.Member.Email || ra.Member.Title;
                                    const statusUser = checkedUserMap.get(key);
                                    if (statusUser && statusUser.OrphanStatus === 'Deleted') {
                                        foundOrphans.push({
                                            key: `item-${item.Id}-${ra.PrincipalId}`,
                                            listName: list.Title,
                                            listId: list.Id,
                                            itemId: item.Id,
                                            // Use Title (PermissionService maps FileLeafRef to Title for files)
                                            itemName: item.Title || "Untitled Item",
                                            itemUrl: item.ServerRelativeUrl,
                                            scope: item.FileSystemObjectType === 1 ? 'Folder' : 'Item',
                                            user: { ...ra.Member, OrphanStatus: statusUser.OrphanStatus } as IUser,
                                            roles: ra.RoleDefinitionBindings.map(r => r.Name),
                                            status: statusUser.OrphanStatus
                                        });
                                    }
                                }
                            });
                        });
                    }

                } catch (e) {
                    console.error(`Error scanning list ${list.Title}`, e);
                    if (e instanceof Error && e.message === 'Scan aborted by user.') throw e;
                }
            }

            setProgress(1);
            setProgressLabel('Scan Complete.');
        } catch (e) {
            if (e instanceof Error && e.message === 'Scan aborted by user.') {
                setProgressLabel('Scan Stopped.');
            } else {
                console.error("Scan failed", e);
                setProgressLabel('Scan Failed.');
            }
        } finally {
            setIsScanning(false);
            abortController.current = null;

            // Build Groups Logic (moved here to ensure it runs even on stop)
            const newGroups: IGroup[] = [];
            let currentGroup: IGroup | null = null;

            foundOrphans.forEach((item, index) => {
                if (!currentGroup || currentGroup.name !== item.listName) {
                    if (currentGroup) {
                        const group = currentGroup as IGroup;
                        group.count = index - group.startIndex;
                        newGroups.push(group);
                    }
                    currentGroup = {
                        key: item.listName,
                        name: item.listName,
                        startIndex: index,
                        count: 0,
                        level: 0,
                        isCollapsed: false
                    };
                }
            });
            if (currentGroup) {
                const group = currentGroup as IGroup;
                group.count = foundOrphans.length - group.startIndex;
                newGroups.push(group);
            }

            setOrphans(foundOrphans);
            setGroups(newGroups);
        }
    };

    const confirmRemove = (item: IOrphanResult) => {
        setItemToRemove(item);
        setHideDialog(false);
    };

    const handleRemove = async () => {
        if (!itemToRemove) return;
        setIsRemoving(true);
        let success = false;

        try {
            if (itemToRemove.scope === 'List') {
                success = await permissionService.removeListPermission(itemToRemove.listId, itemToRemove.user.Id);
            } else {
                success = await permissionService.removeItemPermission(itemToRemove.listId, itemToRemove.itemId!, itemToRemove.user.Id);
            }

            if (success) {
                // Remove from local state
                const newOrphans = orphans.filter(o => o.key !== itemToRemove.key);
                setOrphans(newOrphans);

                // Re-calc groups? Or just let them be slightly off until re-scan?
                // Re-calculating groups is safer for UI
                // Simple lazy way: just decrement count of that group? 
                // Better: filter list and re-run group logic.

                // Re-run group logic on filtered list
                const newGroups: IGroup[] = [];
                let currentGroup: IGroup | null = null;
                newOrphans.forEach((item, index) => {
                    if (!currentGroup || currentGroup.name !== item.listName) {
                        if (currentGroup) {
                            const group = currentGroup as IGroup;
                            group.count = index - group.startIndex;
                            newGroups.push(group);
                        }
                        currentGroup = {
                            key: item.listName,
                            name: item.listName,
                            startIndex: index,
                            count: 0,
                            level: 0,
                            isCollapsed: false
                        };
                    }
                });
                if (currentGroup) {
                    const group = currentGroup as IGroup;
                    group.count = newOrphans.length - group.startIndex;
                    newGroups.push(group);
                }
                setGroups(newGroups);

                setRemovalMessage(`Successfully removed permission for ${itemToRemove.user.Title}.`);
            } else {
                setRemovalMessage(`Failed to remove permission. Please try again.`);
            }
        } catch (e) {
            console.error("Error removing permission", e);
            const errorMessage = e instanceof Error ? e.message : 'Unknown error';
            setRemovalMessage(`Error removing permission: ${errorMessage}`);
        } finally {
            setIsRemoving(false);
            setHideDialog(true);
            setItemToRemove(null);
            setTimeout(() => setRemovalMessage(null), 5000);
        }
    };

    const columns: IColumn[] = [
        { key: 'scope', name: 'Scope', fieldName: 'scope', minWidth: 50, maxWidth: 70 },
        {
            key: 'item', name: 'Location', fieldName: 'itemName', minWidth: 200, maxWidth: 300,
            onRender: (item: IOrphanResult) => (
                <Stack>
                    <span style={{ fontWeight: 600 }}>{item.itemName}</span>
                    <span style={{ fontSize: '12px', color: '#666' }}>{item.itemUrl}</span>
                </Stack>
            )
        },
        {
            key: 'user', name: 'Orphan User', fieldName: 'user', minWidth: 150, maxWidth: 200,
            onRender: (item: IOrphanResult) => <UserPersona user={item.user} />
        },
        {
            key: 'status', name: 'Status', fieldName: 'status', minWidth: 80, maxWidth: 100,
            onRender: (item: IOrphanResult) => (
                <span style={{ color: '#a80000', fontWeight: 600 }}>{item.status}</span>
            )
        },
        {
            key: 'actions', name: 'Action', minWidth: 100, maxWidth: 100,
            onRender: (item: IOrphanResult) => (
                <PrimaryButton
                    text="Remove"
                    onClick={() => confirmRemove(item)}
                    styles={{ root: { height: 24, padding: '0 8px', backgroundColor: '#a80000', borderColor: '#a80000' } }}
                />
            )
        }
    ];

    return (
        <div style={{ padding: '20px', fontSize: contentFontSize }}>
            <Stack tokens={{ childrenGap: 20 }}>
                {/* Header Section */}
                <Stack>
                    <h2 style={{ marginTop: 0, marginBottom: 10 }}>Deep Clean: Orphan Detection</h2>
                    <p style={{ marginTop: 0, color: '#605e5c' }}>
                        This tool scans all Lists, Libraries, and unique Item permissions to find deleted or disabled users who still have access.
                    </p>
                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                            text={isScanning ? "Scanning..." : "Start Deep Clean Scan"}
                            onClick={runDeepClean}
                            disabled={isScanning}
                            iconProps={{ iconName: 'Broom' }}
                            styles={{ root: { maxWidth: 220 } }}
                        />
                        {isScanning && (
                            <DefaultButton
                                text="Stop Scan"
                                onClick={stopScan}
                                iconProps={{ iconName: 'Cancel' }}
                                styles={{ root: { color: '#a80000', borderColor: '#a80000' } }}
                            />
                        )}
                    </Stack>
                </Stack>

                {/* Progress */}
                {isScanning && (
                    <ProgressIndicator label={progressLabel} percentComplete={progress} />
                )}

                {/* Messages */}
                {removalMessage && (
                    <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setRemovalMessage(null)}>
                        {removalMessage}
                    </MessageBar>
                )}

                {/* Results Table */}
                {!isScanning && orphans.length > 0 && (
                    <div>
                        <div style={{ marginBottom: 10, fontWeight: 600 }}>Found {orphans.length} orphan permissions</div>
                        <DetailsList
                            items={orphans}
                            groups={groups}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                            groupProps={{
                                showEmptyGroups: true,
                            }}
                        />
                    </div>
                )}

                {!isScanning && progress === 1 && orphans.length === 0 && (
                    <div style={{ padding: 40, textAlign: 'center', color: '#107c10' }}>
                        <div style={{ fontSize: 40, marginBottom: 10 }}>✅</div>
                        No orphans found! Your site content appears clean.
                    </div>
                )}
            </Stack>

            <Dialog
                hidden={hideDialog}
                onDismiss={() => setHideDialog(true)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Confirm Removal',
                    subText: `Are you sure you want to remove permissions for "${itemToRemove?.user.Title}" on "${itemToRemove?.itemName}"?`
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={handleRemove} text={isRemoving ? "Removing..." : "Remove"} disabled={isRemoving} />
                    <DefaultButton onClick={() => setHideDialog(true)} text="Cancel" disabled={isRemoving} />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
