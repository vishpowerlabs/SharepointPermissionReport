import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PermissionServiceImpl } from '../services/PermissionService';
import { MockPermissionService } from '../services/MockPermissionService';
import { IPermissionService } from '../services/IPermissionService';
import { Header } from './Header';
import { StatsCards } from './StatsCards';
import { SitePermissions } from './SitePermissions';
import { ListPermissions } from './ListPermissions';
import { LoadingState } from './LoadingState';
import { DeepScanDialog } from './DeepScanDialog';
import { CheckAccess } from './CheckAccess';
import { SiteAdmins } from './SiteAdmins';
import { SiteGroups } from './SiteGroups';
import { IItemPermission, IRoleAssignment, IListInfo, ISiteStats, IUser, IGroup } from '../models/IPermissionData';
import { exportSitePermissions, exportListPermissions, exportDeepScanResults } from '../utils/CsvExport';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react';
import styles from './PermissionViewer.module.scss';

export interface IPermissionViewerProps {
    spHttpClient: SPHttpClient;
    webUrl: string;
    themeVariant: IReadonlyTheme | undefined;
    headerOpacity?: number;
    showStats?: boolean;
    excludedLists?: string[];

    buttonFontSize?: string;
    showComponentHeader?: boolean;
    webPartTitle?: string;
    webPartTitleFontSize?: string;
    contentFontSize?: string;
    simulateAccessDenied?: boolean;
    useMockData?: boolean;
}

const PermissionViewer: React.FunctionComponent<IPermissionViewerProps> = (props) => {
    const [activeTab, setActiveTab] = React.useState<string>('site');
    const [sitePermissions, setSitePermissions] = React.useState<IRoleAssignment[]>([]);
    const [filteredSitePermissions, setFilteredSitePermissions] = React.useState<IRoleAssignment[]>([]);

    const [lists, setLists] = React.useState<IListInfo[]>([]);
    const [filteredLists, setFilteredLists] = React.useState<IListInfo[]>([]);

    const [stats, setStats] = React.useState<ISiteStats>({ totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 });
    const [isLoading, setIsLoading] = React.useState<boolean>(true);
    const [loadingMessage, setLoadingMessage] = React.useState<string>('Loading site permissions...');
    const [permissionService, setPermissionService] = React.useState<IPermissionService>();
    const [isExporting, setIsExporting] = React.useState<boolean>(false);

    // Site Admins State
    const [siteAdmins, setSiteAdmins] = React.useState<IUser[]>([]);
    const [isLoadingAdmins, setIsLoadingAdmins] = React.useState<boolean>(false);

    // Site Groups State
    const [siteGroups, setSiteGroups] = React.useState<IGroup[]>([]);
    const [isLoadingGroups, setIsLoadingGroups] = React.useState<boolean>(false);

    const [isScanning, setIsScanning] = React.useState<boolean>(false);

    // Deep Scan State
    const [isDeepScanOpen, setIsDeepScanOpen] = React.useState<boolean>(false);
    const [deepScanItems, setDeepScanItems] = React.useState<IItemPermission[]>([]);
    const [deepScanListTitle, setDeepScanListTitle] = React.useState<string>('');
    const [confirmScanList, setConfirmScanList] = React.useState<{ id: string, title: string } | null>(null);
    const [scanNoResults, setScanNoResults] = React.useState<boolean>(false);
    const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

    // Delete Confirmation State
    const [deleteConfirmState, setDeleteConfirmState] = React.useState<{
        isOpen: boolean;
        title: string;
        subText: string;
        onConfirm: () => void;
        onCancel?: () => void;
    }>({ isOpen: false, title: '', subText: '', onConfirm: () => { } });

    // Access Control State
    const [hasAccess, setHasAccess] = React.useState<boolean | null>(null); // null = checking
    const [accessContacts, setAccessContacts] = React.useState<IUser[]>([]);

    React.useEffect(() => {
        let service: IPermissionService;
        if (props.useMockData) {
            console.log("Using Mock Data");
            service = new MockPermissionService();
        } else {
            service = new PermissionServiceImpl(props.spHttpClient, props.webUrl);
        }
        setPermissionService(service);
        checkAccessAndLoad(service);
    }, [props.excludedLists, props.simulateAccessDenied, props.useMockData]);





    // Demo Mode State
    const [isDemoMode, setIsDemoMode] = React.useState<boolean>(false);
    const [demoStep, setDemoStep] = React.useState<number>(0);
    const [highlightStats, setHighlightStats] = React.useState<boolean>(false);
    const [forcedExpandedListId, setForcedExpandedListId] = React.useState<string | null>(null);

    // Header Handlers
    const handlePlayDemo = () => {
        if (isDemoMode) return;
        setIsDemoMode(true);
        setDemoStep(0);

        // Reset view
        setActiveTab('site');
        setHighlightStats(false);
        setForcedExpandedListId(null);
        // setSearchQuery(''); // Need to expose setSearchQuery or handle Check Access search differently?
        // Actually CheckAccess has its own state. We might need a ref or lift state. 
        // Let's keep it simple: We'll switch tabs. Simulating internal component state is hard without refactoring.
    };

    React.useEffect(() => {
        if (!isDemoMode) return;

        // Steps Timing
        let timer: any;

        const runStep = () => {
            switch (demoStep) {
                case 0: // Start - Highlight Stats
                    setHighlightStats(true);
                    timer = setTimeout(() => setDemoStep(1), 1500);
                    break;
                case 1: // Move to Site Permissions
                    setHighlightStats(false);
                    setActiveTab('site');
                    timer = setTimeout(() => setDemoStep(2), 2000); // Let them see it
                    break;
                case 2: // Move to Lists
                    setActiveTab('lists');
                    timer = setTimeout(() => setDemoStep(3), 1000); // Short pause before action
                    break;
                case 3: // Expand a List
                    {
                        // Find a list with unique perms to expand
                        const uniqueList = lists.find(l => l.HasUniqueRoleAssignments);
                        if (uniqueList) {
                            setForcedExpandedListId(uniqueList.Id);
                        }
                        timer = setTimeout(() => setDemoStep(4), 2500); // Wait while expanded
                    }
                    break;
                case 4: // Close List / Move to Groups
                    setForcedExpandedListId(null);
                    setActiveTab('groups');
                    timer = setTimeout(() => setDemoStep(5), 2000);
                    break;
                case 5: // Move to Site Admins
                    setActiveTab('admins');
                    timer = setTimeout(() => setDemoStep(6), 2000);
                    break;
                case 6: // Move to Check Access
                    setActiveTab('check_access');
                    // Note: Can't easily script the internal state of CheckAccess without refactoring.
                    // We will just show the tab.
                    timer = setTimeout(() => setDemoStep(7), 2000);
                    break;
                case 7: // End
                    setIsDemoMode(false);
                    setDemoStep(0);
                    break;
            }
        };

        runStep();

        return () => clearTimeout(timer);
    }, [isDemoMode, demoStep, lists]);

    const checkAccessAndLoad = async (service: IPermissionService) => {
        setIsLoading(true);
        setLoadingMessage('Checking permissions...');

        try {
            const currentUser = await service.getCurrentUser();
            let isAllowed = currentUser.IsSiteAdmin;

            // If not site admin, check owners group
            if (!isAllowed) {
                const owners = await service.getSiteOwners();
                const isOwner = owners.some(o => o.LoginName === currentUser.LoginName || o.Email === currentUser.Email);
                if (isOwner) {
                    isAllowed = true;
                }
            }

            // SIMULATION OVERRIDE
            if (props.simulateAccessDenied) {
                console.log("Simulating Access Denied via Web Part Property");
                isAllowed = false;
            }

            if (isAllowed) {
                setHasAccess(true);
                loadData(service);
            } else {
                // Access Denied: Load Admins and Owners for contact info
                setLoadingMessage('Loading contact information...');
                const admins = await service.getSiteAdmins();
                const owners = await service.getSiteOwners();

                // Combine and merge roles
                const mergedContacts = new Map<string, IUser>();

                // Helper to add/merge
                const addContact = (u: IUser) => {
                    const key = u.LoginName || u.Email || u.Title;
                    if (mergedContacts.has(key)) {
                        const existing = mergedContacts.get(key)!;
                        // Merge flags
                        existing.IsSiteAdmin = existing.IsSiteAdmin || u.IsSiteAdmin;
                        existing.IsSiteOwner = existing.IsSiteOwner || u.IsSiteOwner;
                    } else {
                        mergedContacts.set(key, { ...u });
                    }
                };

                admins.forEach(addContact);
                owners.forEach(addContact);

                const uniqueContacts = Array.from(mergedContacts.values());

                setAccessContacts(uniqueContacts);
                setHasAccess(false);
                setIsLoading(false);
            }

        } catch (error) {
            console.error("Error checking access", error);
            // Default to denied on error for security
            setHasAccess(false);
            setIsLoading(false);
        }
    };

    const loadData = async (service: IPermissionService) => {
        // isLoading is already true from checkAccessAndLoad
        setLoadingMessage('Loading site permissions...');

        // Load site permissions
        const sitePerms = await service.getSiteRoleAssignments();
        setSitePermissions(sitePerms);
        setFilteredSitePermissions(sitePerms);

        // Load stats
        const siteStats = await service.getSiteStats();

        // Load lists
        setLoadingMessage('Loading lists and libraries...');
        const listsData = await service.getLists(props.excludedLists);

        // Update stats with unique permissions count from lists
        const uniqueListsCount = listsData.filter(l => l.HasUniqueRoleAssignments).length;

        setStats({
            ...siteStats,
            uniquePermissionsCount: uniqueListsCount + (sitePerms.length > 0 ? 1 : 0)
        });

        // Load Site Admins
        setIsLoadingAdmins(true);
        try {
            const admins = await service.getSiteAdmins();
            setSiteAdmins(admins);
        } catch (e) {
            console.error("Error loading admins", e);
        } finally {
            setIsLoadingAdmins(false);
        }

        // Load Site Groups
        setIsLoadingGroups(true);
        try {
            const groups = await service.getSiteGroups();
            setSiteGroups(groups);
        } catch (e) {
            console.error("Error loading groups", e);
        } finally {
            setIsLoadingGroups(false);
        }

        // Sort lists: Unique first, then Inherited
        listsData.sort((a, b) => {
            if (a.HasUniqueRoleAssignments === b.HasUniqueRoleAssignments) return 0;
            return a.HasUniqueRoleAssignments ? -1 : 1;
        });

        setLists(listsData);
        setFilteredLists(listsData);

        setIsLoading(false);
    };

    // ... (existing handlers)

    if (hasAccess === false) {
        return (
            <div className={styles.permissionViewer}>
                <div className={styles.webpartContainer}>
                    <Header
                        onRefresh={() => checkAccessAndLoad(permissionService!)}
                        isLoading={isLoading}
                        themeVariant={props.themeVariant}
                        opacity={props.headerOpacity ?? 100}
                        title={props.webPartTitle}
                        titleFontSize={props.webPartTitleFontSize}
                    />
                    <div style={{ padding: '20px', textAlign: 'center' }}>
                        <h2 style={{ color: '#d13438' }}>You don't have access to this report</h2>
                        <p style={{ marginBottom: '20px' }}>Please check with the Site Owners or Site Administrators listed below.</p>

                        <SiteAdmins users={accessContacts} isLoading={false} />
                    </div>
                </div>
            </div>
        );
    }


    const handleRefresh = () => {
        if (permissionService) {
            loadData(permissionService);
        }
    };

    /* onSearch removed as unused */

    const handleGetListPermissions = async (listId: string) => {
        if (!permissionService) return [];
        // Find list title
        const list = lists.find(l => l.Id === listId);
        return await permissionService.getListRoleAssignments(listId, list ? list.Title : '');
    };

    const handleRemoveSitePermission = (principalId: number, principalName: string) => {
        setDeleteConfirmState({
            isOpen: true,
            title: `Remove Permissions?`,
            subText: `Are you sure you want to remove permissions for ${principalName || 'this user'}? This will remove all permissions for this user on this site.`,
            onConfirm: () => executeRemoveSitePermission(principalId),
            onCancel: () => { /* No-op for site perms */ }
        });
    };

    const executeRemoveSitePermission = async (principalId: number) => {
        if (!permissionService) return;
        setDeleteConfirmState(prev => ({ ...prev, isOpen: false }));
        setIsLoading(true);
        try {
            const success = await permissionService.removeSitePermission(principalId);
            if (success) {
                const sitePerms = await permissionService.getSiteRoleAssignments();
                setSitePermissions(sitePerms);
                setFilteredSitePermissions(sitePerms);
            } else {
                setErrorMessage("Failed to remove permission. Please try again or check console for details.");
            }
        } catch (error) {
            console.error("Error removing site permission", error);
            setErrorMessage("An unexpected error occurred while removing permission.");
        } finally {
            setIsLoading(false);
        }
    };

    const handleRemoveListPermission = async (listId: string, principalId: number, principalName: string): Promise<boolean> => {
        return new Promise<boolean>((resolve) => {
            setDeleteConfirmState({
                isOpen: true,
                title: `Remove Permissions?`,
                subText: `Are you sure you want to remove permissions for ${principalName || 'this user'} on this list?`,
                onConfirm: async () => {
                    setDeleteConfirmState(prev => ({ ...prev, isOpen: false }));
                    if (!permissionService) { resolve(false); return; }
                    try {
                        const success = await permissionService.removeListPermission(listId, principalId);
                        if (!success) {
                            setErrorMessage("Failed to remove list permission. Please try again or check console for details.");
                        }
                        resolve(success);
                    } catch (error) {
                        console.error("Error removing list permission", error);
                        setErrorMessage("An unexpected error occurred while removing list permission.");
                        resolve(false);
                    }
                },
                onCancel: () => {
                    resolve(false);
                }
            });
        });
    };

    const handleRemoveDeepScanItemPermission = (itemId: number, principalId: number, principalName: string) => {
        setDeleteConfirmState({
            isOpen: true,
            title: `Remove Permissions?`,
            subText: `Are you sure you want to remove permissions for ${principalName || 'this user'} on this item?`,
            onConfirm: () => executeRemoveDeepScanItemPermission(itemId, principalId),
            onCancel: () => { /* No-op */ }
        });
    };

    const executeRemoveDeepScanItemPermission = async (itemId: number, principalId: number) => {
        if (!permissionService || !deepScanListTitle) return;
        setDeleteConfirmState(prev => ({ ...prev, isOpen: false }));

        const list = lists.find(l => l.Title === deepScanListTitle);
        if (!list) return;

        try {
            const success = await permissionService.removeItemPermission(list.Id, itemId, principalId);
            if (success) {
                setDeepScanItems(prevItems => {
                    return prevItems.map(item => {
                        if (item.Id === itemId) {
                            const newRoles = item.RoleAssignments.filter(ra => ra.PrincipalId !== principalId);
                            if (newRoles.length === 0) {
                                return { ...item, RoleAssignments: [] };
                            }
                            return { ...item, RoleAssignments: newRoles };
                        }
                        return item;
                    }).filter(item => item !== null) as IItemPermission[];
                });
            } else {
                setErrorMessage("Failed to remove item permission.");
            }
        } catch (error) {
            console.error("Error removing item permission", error);
            setErrorMessage("An unexpected error occurred while removing item permission.");
        }
    };

    const handleRemoveFromGroup = (groupId: number, userId: number, userName: string) => {
        setDeleteConfirmState({
            isOpen: true,
            title: `Remove from Group?`,
            subText: `Are you sure you want to remove '${userName}' from this group?`,
            onConfirm: () => executeRemoveFromGroup(groupId, userId),
            onCancel: () => { /* No-op */ }
        });
    };

    const executeRemoveFromGroup = async (groupId: number, userId: number) => {
        if (!permissionService) return;
        setDeleteConfirmState(prev => ({ ...prev, isOpen: false }));
        setIsLoading(true);

        try {
            const success = await permissionService.removeUserFromGroup(groupId, userId);
            if (success) {
                // We need to refresh the Site Permissions view if a group was expanded. 
                // Since SitePermissions manages its own 'groupMembers' state but we don't have access to it easily here...
                // Ideally, we should lift that state up or force a refresh.
                // Re-loading the permissions effectively resets the view which is acceptable but collapses groups.
                // A better UX would be to just reload the data.

                // Reload data to ensure consistency
                await loadData(permissionService);
            } else {
                setErrorMessage("Failed to remove user from group.");
            }
        } catch (error) {
            console.error("Error removing user from group", error);
            setErrorMessage("An unexpected error occurred.");
        } finally {
            setIsLoading(false);
        }
    };

    const handleExport = async () => {
        if (!permissionService) return;
        setIsExporting(true);

        try {
            if (activeTab === 'site') {
                await exportSitePermissions(filteredSitePermissions, permissionService);
            } else {
                await exportListPermissions(filteredLists, permissionService);
            }
        } catch (error) {
            console.error("Export failed", error);
            alert("Export failed. See console for details.");
        } finally {
            setIsExporting(false);
        }
    };


    // ...

    const handleDeepScan = async (listId: string) => {
        const list = lists.find(l => l.Id === listId);
        if (!list) return;
        setScanNoResults(false);
        setConfirmScanList({ id: list.Id, title: list.Title });
    };

    const executeDeepScan = async () => {
        if (!permissionService || !confirmScanList) return;

        const listId = confirmScanList.id;
        const listTitle = confirmScanList.title;

        setIsScanning(true);
        // setIsExporting(true); // Don't use export flag for scan loading
        try {
            const items = await permissionService.getUniquePermissionItems(listId);

            if (items.length > 0) {
                setDeepScanItems(items);
                setDeepScanListTitle(listTitle);
                setIsDeepScanOpen(true);
                setConfirmScanList(null); // Close confirm dialog only if opening results
            } else {
                setScanNoResults(true);
                // Keep dialog open to show message
            }
        } catch (e) {
            console.error(e);
            alert("Error during deep scan.");
            setConfirmScanList(null); // Close on error too
        } finally {
            setIsScanning(false);

        }
    };

    const downloadDeepScanResults = () => {
        exportDeepScanResults(deepScanItems, deepScanListTitle);
    };

    const getDialogTitle = (): string => {
        if (scanNoResults) return 'Deep Scan Complete';
        if (isScanning) return 'Deep Scan in Progress...';
        return 'Start Deep Scan?';
    };

    const getDialogSubText = (): string => {
        if (scanNoResults) {
            return `No items with unique permissions were found in "${confirmScanList?.title}". All items inherit permissions.`;
        }
        if (isScanning) {
            return `Scanning "${confirmScanList?.title}". This may take a few moments depending on the number of items...`;
        }
        return `This will verify every single item in "${confirmScanList?.title}" to find unique permissions. This might take a while for large lists. Continue?`;
    };

    return (
        <div className={styles.permissionViewer}>
            <div className={styles.webpartContainer}>
                {(props.showComponentHeader !== false) && (
                    <Header
                        title={props.webPartTitle}
                        titleFontSize={props.webPartTitleFontSize}
                        themeVariant={props.themeVariant}
                        opacity={props.headerOpacity}
                        isLoading={isLoading}
                        onRefresh={() => checkAccessAndLoad(permissionService!)}
                        onPlayDemo={handlePlayDemo}
                        isDemoRunning={isDemoMode}
                    />
                )}

                {(props.showStats !== false) && <StatsCards stats={stats} highlight={highlightStats} />}

                <div className={styles.tabsContainer}>
                    <Pivot
                        onLinkClick={(item) => {
                            if (item?.props.itemKey) {
                                setActiveTab(item.props.itemKey);
                            }
                        }}
                        selectedKey={activeTab}
                    >
                        <PivotItem headerText="Site Permissions" itemKey="site" itemIcon="Shield" />
                        <PivotItem headerText="Lists & Libraries" itemKey="lists" itemIcon="List" />
                        <PivotItem headerText="Groups" itemKey="groups" itemIcon="Group" />
                        <PivotItem headerText="Check Access" itemKey="check_access" itemIcon="UserOptional" />
                        <PivotItem headerText="Site Admins" itemKey="admins" itemIcon="Admin" />
                    </Pivot>
                </div>

                <div className={styles.toolbar}>
                    {(activeTab === 'site' || activeTab === 'lists') && (
                        <DefaultButton
                            text={isExporting ? "Exporting..." : "Export to CSV"}
                            iconProps={{ iconName: 'Download' }}
                            onClick={handleExport}
                            disabled={isExporting || isLoading}
                            className={styles.exportBtn}
                            styles={{
                                root: { height: '32px' }, // Maintain height
                                label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                            }}
                        />
                    )}
                </div>

                <div className={styles.content}>
                    {isLoading && <LoadingState message={loadingMessage} />}

                    {!isLoading && activeTab === 'site' && (
                        <SitePermissions
                            permissions={filteredSitePermissions}
                            permissionService={permissionService}
                            contentFontSize={props.contentFontSize}
                            onRemovePermission={handleRemoveSitePermission}
                            onRemoveFromGroup={handleRemoveFromGroup}
                        />
                    )}

                    {!isLoading && activeTab === 'lists' && (
                        <ListPermissions
                            lists={filteredLists}
                            getListPermissions={handleGetListPermissions}
                            onScanItems={handleDeepScan}
                            themeVariant={props.themeVariant}
                            buttonFontSize={props.buttonFontSize}
                            contentFontSize={props.contentFontSize}
                            onRemovePermission={handleRemoveListPermission}
                            forcedExpandedListId={forcedExpandedListId}
                        />
                    )}

                    {!isLoading && activeTab === 'groups' && (
                        <SiteGroups
                            groups={siteGroups}
                            isLoading={isLoadingGroups}
                            permissionService={permissionService!}
                            contentFontSize={props.contentFontSize}
                        />
                    )}

                    {!isLoading && activeTab === 'check_access' && (
                        <CheckAccess
                            permissionService={permissionService!}
                            sitePermissions={sitePermissions}
                            lists={lists}
                            contentFontSize={props.contentFontSize}
                        />
                    )}

                    {!isLoading && activeTab === 'admins' && (
                        <SiteAdmins
                            users={siteAdmins}
                            isLoading={isLoadingAdmins}
                        />
                    )}
                </div>
            </div>

            <DeepScanDialog
                isOpen={isDeepScanOpen}
                onDismiss={() => setIsDeepScanOpen(false)}
                listTitle={deepScanListTitle}
                items={deepScanItems}
                onDownload={downloadDeepScanResults}
                buttonFontSize={props.buttonFontSize}
                contentFontSize={props.contentFontSize}
                onRemovePermission={handleRemoveDeepScanItemPermission}
            />

            <Dialog
                hidden={!confirmScanList}
                onDismiss={() => { if (!isScanning) setConfirmScanList(null); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: getDialogTitle(),
                    subText: getDialogSubText()
                }}
            >
                {isScanning ? (
                    <Spinner size={SpinnerSize.large} label="Scanning for unique permissions..." />
                ) : (
                    <DialogFooter>
                        {scanNoResults ? (
                            <PrimaryButton
                                onClick={() => setConfirmScanList(null)}
                                text="Close"
                                styles={{
                                    root: { height: '32px' },
                                    label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                                }}
                            />
                        ) : (
                            <>
                                <PrimaryButton
                                    onClick={executeDeepScan}
                                    text="Start Scan"
                                    styles={{
                                        root: { height: '32px' },
                                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                                    }}
                                />
                                <DefaultButton
                                    onClick={() => setConfirmScanList(null)}
                                    text="Cancel"
                                    styles={{
                                        root: { height: '32px' },
                                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                                    }}
                                />
                            </>
                        )}
                    </DialogFooter>
                )}
            </Dialog>

            {/* Delete Confirmation Dialog */}
            <Dialog
                hidden={!deleteConfirmState.isOpen}
                onDismiss={() => setDeleteConfirmState(prev => ({ ...prev, isOpen: false }))}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: deleteConfirmState.title,
                    subText: deleteConfirmState.subText,
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 450 } }
                }}
            >
                <DialogFooter>
                    <PrimaryButton
                        onClick={deleteConfirmState.onConfirm}
                        text="Remove"
                        styles={{
                            root: { background: '#d13438', border: '1px solid #d13438' }, // Red color for danger
                            rootHovered: { background: '#a4262c' },
                            label: { fontWeight: 600 }
                        }}
                    />
                    <DefaultButton
                        onClick={() => setDeleteConfirmState(prev => ({ ...prev, isOpen: false }))}
                        text="Cancel"
                    />
                </DialogFooter>
            </Dialog>

            {/* Error Dialog */}
            <Dialog
                hidden={!errorMessage}
                onDismiss={() => setErrorMessage(null)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Error',
                    subText: errorMessage || 'An unexpected error occurred.'
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={() => setErrorMessage(null)} text="OK" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};

export default PermissionViewer;
