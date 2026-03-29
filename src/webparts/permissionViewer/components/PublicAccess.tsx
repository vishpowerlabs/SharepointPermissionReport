import * as React from 'react';
import { PrimaryButton, DefaultButton, MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, Stack, Checkbox, TextField, Label } from '@fluentui/react';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode, IGroup } from '@fluentui/react/lib/DetailsList';
import { IPermissionService } from '../services/IPermissionService';
import { IListInfo, IPublicAccessResult, IRoleAssignment } from '../models/IPermissionData';
import { exportPublicAccessResults } from '../utils/CsvExport';
import { LoadingState } from './LoadingState';

export interface IPublicAccessProps {
    permissionService: IPermissionService;
    lists: IListInfo[];
    contentFontSize?: string;
}

export const PublicAccess: React.FunctionComponent<IPublicAccessProps> = (props) => {
    const { permissionService, lists, contentFontSize } = props;

    const [isScanning, setIsScanning] = React.useState(false);
    const [progress, setProgress] = React.useState<number>(0);
    const [progressLabel, setProgressLabel] = React.useState<string>('');
    const [results, setResults] = React.useState<IPublicAccessResult[]>([]);
    const [groups, setGroups] = React.useState<IGroup[]>([]);

    // Removal State
    const [itemToRemove, setItemToRemove] = React.useState<IPublicAccessResult | null>(null);
    const [isRemoving, setIsRemoving] = React.useState(false);
    const [hideDialog, setHideDialog] = React.useState(true);
    const [removalMessage, setRemovalMessage] = React.useState<string | null>(null);
    const [scanDuration, setScanDuration] = React.useState<string | null>(null);
    const startTimeRef = React.useRef<number>(0);

    const [publicGroupIds, setPublicGroupIds] = React.useState<Set<number>>(new Set());
    const [showAllItems, setShowAllItems] = React.useState(false);

    // Debug Item State
    const [debugQuery, setDebugQuery] = React.useState<string>("");
    const [debugResult, setDebugResult] = React.useState<string>("");

    // Abort Controller for stopping scan
    const abortController = React.useRef<AbortController | null>(null);

    const onDebugSearch = async () => {
        if (!debugQuery) return;
        setDebugResult("Searching...");
        try {
            // Search in the first library found (usually Documents)
            // In a real scenario we might want to pick the library, but for now assuming user knows where it is.
            // Or better: Search across all lists? That's heavy.
            // Let's use the first list in the tree for now or "Documents".
            const lists = await permissionService.getLists();
            const docLib = lists.find(l => l.Title === "Documents") || lists[0];

            if (!docLib) {
                setDebugResult("No library found to search.");
                return;
            }

            setDebugResult(`Searching for '${debugQuery}' in '${docLib.Title}'...`);

            // Custom search logic using filter
            // We can't easily search recursively by name via REST without search API (which might be delayed).
            // We'll try a CAML query for FileLeafRef contains.

            // Actually, the Service doesn't expose a generic search.
            // Let's just create a quick fetch here using the context URL provided by service?
            // No, let's add a helper to PermissionService or just use a raw fetch from here?
            // Accessing service private members is hard.
            // Let's assume the component can ask the service to "debugItem".

            // Since I can't easily change the service interface and wait for build in one step without risk,
            // I will use `permissionService` to get ALL items (we just optimized it!) filtering client side? 
            // No, 5000 items is too many to fetch just for debug.

            // Hack: We'll modify PermissionService to add 'findItemByName' if needed, 
            // OR just ask the user to Check the "Show All" and use browser search?
            // They said it's not there.

            // Let's try to fetch just that item by ID if they know it? No they don't.
            // Let's add `findItemPermissionsByName` to IPermissionService.

        } catch (e: any) {
            setDebugResult("Error: " + e.message);
        }
    };

    const isPublicPrincipal = (title: string, loginName: string, principalId: number): string | null => {
        const t = (title || '').toLowerCase();
        const l = (loginName || '').toLowerCase();

        // Check 1: Direct Match
        if (t.indexOf('everyone') !== -1 || l.indexOf('everyone') !== -1 || l.indexOf('c:0(.s|true') !== -1 || l.indexOf('spo-grid-all-users') !== -1) {
            return 'Everyone';
        }
        if (t.indexOf('anonymous') !== -1 || l.indexOf('anonymous') !== -1) {
            return 'Anonymous';
        }
        if (t.indexOf('authenticated users') !== -1 || l.indexOf('nt authority\\authenticated users') !== -1) {
            return 'Authenticated Users';
        }
        if (t.indexOf('domain users') !== -1) {
            return 'Domain Users';
        }

        // Check 2: Sharing Link / Guest Access
        if (t.indexOf('sharing link') !== -1 || t.indexOf('link for') !== -1 || l.indexOf('urn:spo:guest') !== -1) {
            return 'Sharing Link';
        }

        // Check 3: "Guest" keyword in title (Careful, "Guest Contributor" is a role, not a user usually, but sometimes "Guest Users" group exists)
        if (t.indexOf('all users') !== -1 || (t.indexOf('guest') !== -1 && t.indexOf('contributor') === -1)) {
            // Exclude "Guest Contributor" ROLE if it appears in name (unlikely for Principal Name)
            // But include "All Users" or "Guests" group
            return 'Guest/Public Group';
        }

        // Check 3: Nested in Public Group
        if (publicGroupIds.has(principalId)) {
            return 'Public Group Member';
        }

        return null;
    };

    const stopScan = () => {
        if (abortController.current) {
            abortController.current.abort();
        }
    };

    const runScan = async () => {
        setIsScanning(true);
        setResults([]);
        setGroups([]);
        setProgress(0);
        setRemovalMessage(null);
        setScanDuration(null);
        startTimeRef.current = Date.now();

        // Init AbortController
        abortController.current = new AbortController();
        const signal = abortController.current.signal;

        const foundItems: IPublicAccessResult[] = [];
        const uniqueLists = lists.filter(l => l.HasUniqueRoleAssignments).sort((a, b) => a.Title.localeCompare(b.Title));
        const totalSteps = uniqueLists.length;

        try {
            // 0. Pre-scan: Identify Public Groups
            setProgressLabel('Analyzing Site Groups...');
            const siteGroups = await permissionService.getSiteGroups();
            const publicIds = new Set<number>();

            siteGroups.forEach(g => {
                // Check if group contains Everyone/Anonymous
                const hasPublicMember = g.Users?.some(u => {
                    const l = (u.LoginName || '').toLowerCase(); // Fix potential undefined
                    return l.indexOf('everyone') !== -1 ||
                        l.indexOf('c:0(.s|true') !== -1 ||
                        l.indexOf('spo-grid-all-users') !== -1 ||
                        l.indexOf('anonymous') !== -1 ||
                        l.indexOf('domain users') !== -1 ||
                        l.indexOf('urn:spo:guest') !== -1; // Catch guest/sharing links
                });

                if (hasPublicMember) {
                    publicIds.add(g.Id);
                    console.log(`Deep Scan: Found Public Group: ${g.Title} (ID: ${g.Id})`);
                }
            });
            setPublicGroupIds(publicIds);

            for (let i = 0; i < uniqueLists.length; i++) {
                if (signal.aborted) throw new Error('Scan aborted by user.');

                const list = uniqueLists[i];
                setProgress((i / totalSteps));
                setProgressLabel(`Scanning ${list.Title}...`);

                // 1. Check List Level
                try {
                    const listPerms = await permissionService.getListRoleAssignments(list.Id, list.Title);
                    checkPermissions(listPerms, list, 'List', null, foundItems, publicIds, showAllItems);

                    // 2. Check Item Level
                    // Pass signal to service to allow cancellation during fetch
                    const items = await permissionService.getUniquePermissionItems(list.Id, signal);

                    items.forEach(item => {
                        checkPermissions(item.RoleAssignments, list, item.FileSystemObjectType === 1 ? 'Folder' : 'Item', item, foundItems, publicIds, showAllItems);
                    });

                } catch (e) {
                    console.error(`Error scanning list ${list.Title}`, e);
                    // Decide if we should continue or break. Usually continue for other lists unless aborted.
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
            buildGroups(foundItems);
            setResults(foundItems);

            const endTime = Date.now();
            const durationMs = endTime - startTimeRef.current;
            const minutes = Math.floor(durationMs / 60000);
            const seconds = ((durationMs % 60000) / 1000).toFixed(0);
            setScanDuration(`Scan took ${minutes}m ${seconds}s`);
        }
    };

    const checkPermissions = (
        roles: IRoleAssignment[],
        list: IListInfo,
        scope: 'List' | 'Item' | 'Folder',
        item: any | null,
        foundItems: IPublicAccessResult[],
        cmdPublicIds: Set<number>,
        includeAll: boolean
    ) => {
        roles.forEach(r => {
            // Re-implement check to use the set passed in (closure freshness)
            const isPublic = (title: string, loginName: string, id: number): string | null => {
                const t = (title || '').toLowerCase();
                const l = (loginName || '').toLowerCase();

                if (t.indexOf('everyone') !== -1 || l.indexOf('everyone') !== -1 || l.indexOf('c:0(.s|true') !== -1 || l.indexOf('spo-grid-all-users') !== -1) return 'Everyone';
                if (t.indexOf('anonymous') !== -1 || l.indexOf('anonymous') !== -1) return 'Anonymous';
                if (t.indexOf('authenticated users') !== -1 || l.indexOf('nt authority\\authenticated users') !== -1) return 'Authenticated Users';
                if (t.indexOf('domain users') !== -1) return 'Domain Users';
                if (t.indexOf('sharing link') !== -1 || t.indexOf('link for') !== -1 || l.indexOf('urn:spo:guest') !== -1) return 'Sharing Link';
                if (t.indexOf('all users') !== -1 || (t.indexOf('guest') !== -1 && t.indexOf('contributor') === -1)) return 'Guest/Public Group';
                if (cmdPublicIds.has(id)) return 'Public Group Member';
                return null;
            };

            const publicType = isPublic(r.Member.Title, r.Member.LoginName || '', r.PrincipalId);

            if (publicType || includeAll) {
                foundItems.push({
                    key: `${scope}-${list.Id}-${item ? item.Id : 'root'}-${r.PrincipalId}`,
                    listName: list.Title,
                    listId: list.Id,
                    itemId: item ? item.Id : undefined,
                    itemName: item ? (item.Title || item.FileLeafRef || 'Untitled') : list.Title,
                    itemUrl: item ? item.ServerRelativeUrl : list.ServerRelativeUrl,
                    scope: scope,
                    principalName: r.Member.Title,
                    principalType: publicType || 'Private', // Display 'Private' for debug items
                    roles: r.RoleDefinitionBindings.map(d => d.Name)
                });
            }
        });
    };

    const buildGroups = (items: IPublicAccessResult[]) => {
        const newGroups: IGroup[] = [];
        let currentGroup: IGroup | null = null;

        items.forEach((item, index) => {
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
            group.count = items.length - group.startIndex;
            newGroups.push(group);
        }
        setGroups(newGroups);
    };

    const confirmRemove = (item: IPublicAccessResult) => {
        setItemToRemove(item);
        setHideDialog(false);
    };

    const handleRemove = async () => {
        if (!itemToRemove) return;
        setIsRemoving(true);
        let success = false;

        const parts = itemToRemove.key.split('-');
        const principalId = parseInt(parts[parts.length - 1]);

        try {
            if (itemToRemove.scope === 'List') {
                success = await permissionService.removeListPermission(itemToRemove.listId, principalId);
            } else {
                success = await permissionService.removeItemPermission(itemToRemove.listId, itemToRemove.itemId!, principalId);
            }

            if (success) {
                const newResults = results.filter(r => r.key !== itemToRemove.key);
                setResults(newResults);
                buildGroups(newResults);
                setRemovalMessage(`Successfully removed access for ${itemToRemove.principalName}.`);
            } else {
                setRemovalMessage(`Failed to remove permission.`);
            }
        } catch (e: any) {
            console.error("Error removing", e);
            setRemovalMessage("Error removing permission.");
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
            onRender: (item: IPublicAccessResult) => (
                <Stack>
                    <span style={{ fontWeight: 600 }}>{item.itemName}</span>
                    <span style={{ fontSize: '12px', color: '#666' }}>{item.itemUrl}</span>
                </Stack>
            )
        },
        { key: 'principal', name: 'Public Group', fieldName: 'principalName', minWidth: 150, maxWidth: 200 },
        {
            key: 'type', name: 'Type', fieldName: 'principalType', minWidth: 100, maxWidth: 120,
            onRender: (item: IPublicAccessResult) => (
                <span style={{ color: '#d13438', fontWeight: 600, backgroundColor: '#fde7e9', padding: '2px 6px', borderRadius: '4px' }}>
                    {item.principalType}
                </span>
            )
        },
        { key: 'role', name: 'Permission', fieldName: 'roles', minWidth: 100, maxWidth: 150, onRender: (i: IPublicAccessResult) => i.roles.join(', ') },
        {
            key: 'actions', name: 'Action', minWidth: 100, maxWidth: 100,
            onRender: (item: IPublicAccessResult) => (
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
                <Stack>
                    <h2 style={{ marginTop: 0, marginBottom: 10 }}>Public Exposure Check</h2>
                    <p style={{ marginTop: 0, color: '#605e5c' }}>
                        Identify content shared with "Everyone", "Anonymous Users", or "Authenticated Users".
                    </p>
                    <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                        <PrimaryButton
                            text={isScanning ? "Scanning..." : "Start Public Scan"}
                            onClick={runScan}
                            disabled={isScanning}
                            iconProps={{ iconName: 'World' }}
                        />
                        <Checkbox
                            label="Show All Unique Items (Debug)"
                            checked={showAllItems}
                            onChange={(ev, checked) => setShowAllItems(!!checked)}
                            disabled={isScanning}
                        />
                        {isScanning && (
                            <DefaultButton
                                text="Stop Scan"
                                onClick={stopScan}
                                iconProps={{ iconName: 'Cancel' }}
                                styles={{ root: { color: '#a80000', borderColor: '#a80000' } }}
                            />
                        )}
                        {!isScanning && results.length > 0 && (
                            <DefaultButton
                                text="Export Results"
                                onClick={() => exportPublicAccessResults(results)}
                                iconProps={{ iconName: 'Download' }}
                            />
                        )}
                    </Stack>
                </Stack>

                {isScanning && (
                    <LoadingState message={progressLabel} progress={progress} />
                )}

                {scanDuration && !isScanning && (
                    <MessageBar messageBarType={MessageBarType.success} styles={{ root: { marginTop: 10 } }}>
                        {scanDuration}
                    </MessageBar>
                )}

                {removalMessage && (
                    <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setRemovalMessage(null)}>
                        {removalMessage}
                    </MessageBar>
                )}

                {!isScanning && results.length > 0 && (
                    <div>
                        <div style={{ marginBottom: 10, fontWeight: 600 }}>Found {results.length} exposed items</div>
                        <DetailsList
                            items={results}
                            groups={groups}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                            groupProps={{ showEmptyGroups: true }}
                        />
                    </div>
                )}

                {!isScanning && progress === 1 && results.length === 0 && (
                    <div style={{ padding: 40, textAlign: 'center', color: '#107c10' }}>
                        <div style={{ fontSize: 40, marginBottom: 10 }}>✅</div>
                        No public exposure found!
                    </div>
                )}
            </Stack>

            <Dialog
                hidden={hideDialog}
                onDismiss={() => setHideDialog(true)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Remove Public Access',
                    subText: `Are you sure you want to remove "${itemToRemove?.principalName}" from "${itemToRemove?.itemName}"?`
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
