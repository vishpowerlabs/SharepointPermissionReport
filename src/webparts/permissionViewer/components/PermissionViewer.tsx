import * as React from 'react';
import { Nav, INavLinkGroup } from '@fluentui/react/lib/Nav';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PermissionServiceNew } from '../services/PermissionServiceNew';
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
import { SecurityGovernance } from './SecurityGovernance';
import { IItemPermission, IRoleAssignment, IListInfo, ISiteStats, IUser, IGroup, ISiteUsage } from '../models/IPermissionData';
import { exportSitePermissions, exportListPermissions, exportDeepScanResults } from '../utils/CsvExport';

import { Dialog, DialogType, DialogFooter, MessageBar, MessageBarType, Panel, PanelType, Checkbox } from '@fluentui/react';
import styles from './PermissionViewer.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';
import { formatBytes } from '../utils/FormatUtils';

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
    showExternalUserAudit?: boolean;
    showSharingLinks?: boolean;
    showOrphanedUsers?: boolean;
    showSecurityGovernanceTab?: boolean;
    navLayout?: 'left' | 'top';
    storageFormat?: 'Auto' | 'MB' | 'GB' | 'TB';
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

    // Unique Perms Panel State
    const [isUniquePermsPanelOpen, setIsUniquePermsPanelOpen] = React.useState<boolean>(false);

    // Site Details State
    // Site Details used to be here but was unused
    const [siteUsage, setSiteUsage] = React.useState<ISiteUsage | null>(null);

    // Groups Panel State
    const [isGroupsPanelOpen, setIsGroupsPanelOpen] = React.useState<boolean>(false);
    const [expandedGroupId, setExpandedGroupId] = React.useState<number | null>(null);
    const [groupMembers, setGroupMembers] = React.useState<{ [groupId: number]: IUser[] }>({});
    const [showEmptyGroupsOnly, setShowEmptyGroupsOnly] = React.useState<boolean>(false);

    // Storage Panel State
    const [isStoragePanelOpen, setIsStoragePanelOpen] = React.useState<boolean>(false);

    const showConfirmDialog = (title: string, subText: string, onConfirm: () => void, onCancel?: () => void) => {
        setDeleteConfirmState({
            isOpen: true,
            title,
            subText,
            onConfirm,
            onCancel: onCancel || (() => { })
        });
    };

    React.useEffect(() => {
        let service: IPermissionService;
        if (props.useMockData) {
            console.log("Using Mock Data");
            service = new MockPermissionService();
        } else {
            service = new PermissionServiceNew(props.spHttpClient, props.webUrl);
        }
        setPermissionService(service);
        checkAccessAndLoad(service);
    }, [props.excludedLists, props.simulateAccessDenied, props.useMockData]);







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
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
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

        // Load Site Details (Check uniqueness)
        // Load Site Details (Check uniqueness) - Call service but don't store unused state
        await service.getSiteDetails();
        // setSiteDetails(details); -> Unused

        // Load Site Usage
        const usage = await service.getSiteUsage();
        setSiteUsage(usage);

        // Load stats
        const siteStats = await service.getSiteStats();

        // Load lists
        setLoadingMessage('Loading lists and libraries...');
        const listsData = await service.getLists(props.excludedLists);

        // Update stats with unique permissions count from lists
        const uniqueListsCount = listsData.filter(l => l.HasUniqueRoleAssignments).length;

        setStats({
            ...siteStats,
            uniquePermissionsCount: uniqueListsCount
        });

        // Fallback for Storage Usage if API returned 0
        if (usage.storageUsed === 0 && listsData.length > 0) {
            const totalListBytes = listsData.reduce((acc, list) => acc + (list.TotalSize || 0), 0);
            if (totalListBytes > 0) {
                console.log("Using Calculated Storage Fallback:", totalListBytes);
                setSiteUsage({
                    ...usage,
                    storageUsed: totalListBytes
                    // Keep quota as passed from API (likely 0 if failed, but we can't guess quota)
                });
            }
        }

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

    // handleRefresh removed as unused

    // ... (existing handlers)

    if (hasAccess === false) {
        return (
            <div className={styles.permissionViewer}>
                <div className={styles.webpartContainer}>
                    <Header
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




    /* onSearch removed as unused */

    const handleGetListPermissions = async (listId: string) => {
        if (!permissionService) return [];
        // Find list title
        const list = lists.find(l => l.Id === listId);
        return await permissionService.getListRoleAssignments(listId, list ? list.Title : '');
    };

    const handleRemoveSitePermission = (principalId: number, principalName: string) => {
        showConfirmDialog(
            `Remove Permissions?`,
            `Are you sure you want to remove permissions for ${principalName || 'this user'}? This will remove all permissions for this user on this site.`,
            () => { void executeRemoveSitePermission(principalId); }
        );
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

    const executeRemoveListPermission = async (listId: string, principalId: number, resolve: (value: boolean | PromiseLike<boolean>) => void) => {
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
    };

    const showRemoveListPermissionConfirm = (listId: string, principalId: number, principalName: string, resolve: (value: boolean | PromiseLike<boolean>) => void) => {
        showConfirmDialog(
            `Remove Permissions?`,
            `Are you sure you want to remove permissions for ${principalName || 'this user'} on this list?`,
            () => { void executeRemoveListPermission(listId, principalId, resolve); },
            () => { resolve(false); }
        );
    };

    const handleRemoveListPermission = (listId: string, principalId: number, principalName: string): Promise<boolean> => {
        return new Promise<boolean>((resolve) => {
            showRemoveListPermissionConfirm(listId, principalId, principalName, resolve);
        });
    };

    const handleRemoveDeepScanItemPermission = (itemId: number, principalId: number, principalName: string) => {
        showConfirmDialog(
            `Remove Permissions?`,
            `Are you sure you want to remove permissions for ${principalName || 'this user'} on this item?`,
            () => { void executeRemoveDeepScanItemPermission(itemId, principalId); }
        );
    };

    const removePrincipalFromItems = (items: IItemPermission[], itemId: number, principalId: number): IItemPermission[] => {
        return items.map(item => {
            if (item.Id !== itemId) return item;
            const newRoles = item.RoleAssignments.filter(ra => ra.PrincipalId !== principalId);
            return { ...item, RoleAssignments: newRoles };
        });
    };

    const updateDeepScanItemsAfterRemoval = (itemId: number, principalId: number) => {
        setDeepScanItems(prevItems => removePrincipalFromItems(prevItems, itemId, principalId));
    };

    const executeRemoveDeepScanItemPermission = async (itemId: number, principalId: number) => {
        if (!permissionService || !deepScanListTitle) return;
        setDeleteConfirmState(prev => ({ ...prev, isOpen: false }));

        const list = lists.find(l => l.Title === deepScanListTitle);
        if (!list) return;

        try {
            const success = await permissionService.removeItemPermission(list.Id, itemId, principalId);
            if (success) {
                updateDeepScanItemsAfterRemoval(itemId, principalId);
            } else {
                setErrorMessage("Failed to remove item permission.");
            }
        } catch (error) {
            console.error("Error removing item permission", error);
            setErrorMessage("An unexpected error occurred while removing item permission.");
        }
    };

    const handleRemoveFromGroup = (groupId: number, userId: number, userName: string) => {
        showConfirmDialog(
            `Remove from Group?`,
            `Are you sure you want to remove '${userName}' from this group?`,
            () => { void executeRemoveFromGroup(groupId, userId); }
        );
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

    const toggleGroupExpansion = async (groupId: number) => {
        if (expandedGroupId === groupId) {
            setExpandedGroupId(null);
            return;
        }

        setExpandedGroupId(groupId);
        if (!groupMembers[groupId] && permissionService) {
            try {
                const members = await permissionService.getGroupMembers(groupId);
                setGroupMembers(prev => ({ ...prev, [groupId]: members }));
            } catch (error) {
                console.error("Error loading group members", error);
            }
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

    const navGroups: INavLinkGroup[] = [
        {
            links: [
                {
                    name: 'Site Permissions',
                    url: '',
                    key: 'site',
                    icon: 'Shield',
                    onClick: () => { setActiveTab('site'); }
                },
                {
                    name: 'Lists & Libraries',
                    url: '',
                    key: 'lists',
                    icon: 'List',
                    onClick: () => { setActiveTab('lists'); }
                },
                {
                    name: 'Security & Governance',
                    url: '',
                    key: 'governance',
                    icon: 'SecurityGroup',
                    onClick: () => { setActiveTab('governance'); }
                },
                {
                    name: 'Site Groups',
                    url: '',
                    key: 'groups',
                    icon: 'Group',
                    onClick: () => { setActiveTab('groups'); }
                },
                {
                    name: 'Check Access',
                    url: '',
                    key: 'check_access',
                    icon: 'UserOptional',
                    onClick: () => { setActiveTab('check_access'); }
                },
                {
                    name: 'Site Admins',
                    url: '',
                    key: 'admins',
                    icon: 'Admin',
                    onClick: () => { setActiveTab('admins'); }
                }
            ].filter(link => link.key !== 'governance' || props.showSecurityGovernanceTab !== false)
        }
    ];

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

                    />
                )}

                <div className={props.navLayout === 'top' ? styles.layoutContainerTop : styles.layoutContainer}>
                    <div className={props.navLayout === 'top' ? styles.navigationTop : styles.navigation}>
                        {props.navLayout === 'top' ? (
                            <Pivot
                                selectedKey={activeTab}
                                onLinkClick={(item) => {
                                    if (item?.props.itemKey) setActiveTab(item.props.itemKey);
                                }}
                                styles={{
                                    root: { paddingLeft: 16, paddingTop: 8 }
                                }}
                            >
                                <PivotItem headerText="Site Permissions" itemKey="site" itemIcon="Shield" />
                                <PivotItem headerText="Lists & Libraries" itemKey="lists" itemIcon="List" />
                                {(props.showSecurityGovernanceTab !== false) && <PivotItem headerText="Security & Governance" itemKey="governance" itemIcon="SecurityGroup" />}
                                <PivotItem headerText="Site Groups" itemKey="groups" itemIcon="Group" />
                                <PivotItem headerText="Check Access" itemKey="check_access" itemIcon="UserOptional" />
                                <PivotItem headerText="Site Admins" itemKey="admins" itemIcon="Admin" />
                            </Pivot>
                        ) : (
                            <Nav
                                groups={navGroups}
                                selectedKey={activeTab}
                                styles={{
                                    root: {
                                        width: '100%',
                                        height: '100%',
                                        boxSizing: 'border-box',
                                        border: '1px solid transparent',
                                        overflowY: 'auto'
                                    }
                                }}
                            />
                        )}
                    </div>

                    <div className={styles.mainCanvas}>
                        {(props.showStats !== false) && (
                            <StatsCards
                                stats={stats}
                                siteUsage={siteUsage || undefined}
                                highlight={false}
                                onUniquePermissionsClick={() => setIsUniquePermsPanelOpen(true)}
                                onGroupsClick={() => setIsGroupsPanelOpen(true)}
                                onStorageClick={() => setIsStoragePanelOpen(true)}
                                storageFormat={props.storageFormat}
                            />
                        )}

                        <div className={styles.toolbar}>
                            {(activeTab === 'site' || activeTab === 'lists') && (
                                <DefaultButton
                                    text={isExporting ? "Exporting..." : "Export to CSV"}
                                    iconProps={{ iconName: 'Download' }}
                                    onClick={() => { void handleExport(); }}
                                    disabled={isExporting || isLoading}
                                    className={styles.exportBtn}
                                    styles={{
                                        root: { height: '32px' }, // Maintain height
                                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                                    }}
                                />
                            )}
                        </div>

                        {errorMessage && (
                            <MessageBar
                                messageBarType={MessageBarType.error}
                                isMultiline={false}
                                onDismiss={() => setErrorMessage(null)}
                                dismissButtonAriaLabel="Close"
                                styles={{ root: { marginBottom: 10 } }}
                            >
                                {errorMessage}
                            </MessageBar>
                        )}

                        <div className={styles.content}>
                            {isLoading && <LoadingState message={loadingMessage} />}

                            {!isLoading && activeTab === 'site' && (
                                <SitePermissions
                                    permissions={filteredSitePermissions}
                                    permissionService={permissionService}
                                    contentFontSize={props.contentFontSize}
                                    onRemovePermission={handleRemoveSitePermission}
                                    onRemoveFromGroup={handleRemoveFromGroup}
                                    siteGroups={siteGroups}
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
                                    forcedExpandedListId={null}
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

                            {!isLoading && activeTab === 'governance' && (
                                <SecurityGovernance
                                    contentFontSize={props.contentFontSize}
                                    permissionService={permissionService!}
                                    showExternalUserAudit={props.showExternalUserAudit}
                                    showSharingLinks={props.showSharingLinks}
                                    showOrphanedUsers={props.showOrphanedUsers}
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

                    <Panel
                        isOpen={isUniquePermsPanelOpen}
                        onDismiss={() => setIsUniquePermsPanelOpen(false)}
                        isLightDismiss={true}
                        type={PanelType.medium}
                        headerText="Unique Permissions Details"
                        closeButtonAriaLabel="Close"
                    >
                        <p>The following objects have unique permissions (broken inheritance):</p>

                        <div className={styles.permissionTable} style={{ marginTop: 20, border: 'none' }}>
                            {/* Lists */}
                            {lists.filter(l => l.HasUniqueRoleAssignments).map(list => (
                                <div key={list.Id} style={{
                                    padding: '12px 0',
                                    borderBottom: '1px solid #eee',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: 12
                                }}>
                                    <Icon iconName={list.ItemType === 'Library' ? 'DocLibrary' : 'List'} style={{ color: '#0078d4', fontSize: 16 }} />
                                    <div>
                                        <div style={{ fontWeight: 600 }}>{list.Title}</div>
                                        <div style={{ fontSize: 12, color: '#666' }}>{list.ItemType} • {list.ItemCount} items</div>
                                    </div>
                                </div>
                            ))}
                        </div>
                    </Panel>

                    <Panel
                        isOpen={isGroupsPanelOpen}
                        onDismiss={() => {
                            setIsGroupsPanelOpen(false);
                            setExpandedGroupId(null);
                        }}
                        isLightDismiss={true}
                        type={PanelType.medium}
                        headerText="Site Groups"
                        closeButtonAriaLabel="Close"
                    >
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 15 }}>
                            <p style={{ margin: 0 }}>All groups in this site ({siteGroups.length} total):</p>
                            <Checkbox
                                label="Show Empty Groups Only"
                                checked={showEmptyGroupsOnly}
                                onChange={(_, checked) => setShowEmptyGroupsOnly(!!checked)}
                                styles={{ root: { marginBottom: 0 } }}
                            />
                        </div>

                        <div className={styles.permissionTable} style={{ marginTop: 20, border: 'none' }}>
                            {siteGroups
                                .filter(g => !showEmptyGroupsOnly || (g.UserCount === 0 || (g.Users?.length === 0)))
                                .map(group => (
                                    <div key={group.Id} style={{ marginBottom: 10 }}>
                                        <button
                                            style={{
                                                padding: '12px',
                                                border: 'none',
                                                borderBottom: '1px solid #eee',
                                                display: 'flex',
                                                alignItems: 'center',
                                                gap: 12,
                                                cursor: 'pointer',
                                                background: expandedGroupId === group.Id ? '#f3f2f1' : 'transparent',
                                                width: '100%',
                                                textAlign: 'left',
                                                fontFamily: 'inherit'
                                            }}
                                            type="button"
                                            onClick={() => { void toggleGroupExpansion(group.Id); }}
                                            onKeyDown={(e) => {
                                                if (e.key === 'Enter' || e.key === ' ') {
                                                    e.preventDefault();
                                                    void toggleGroupExpansion(group.Id);
                                                }
                                            }}
                                        >
                                            <Icon iconName={expandedGroupId === group.Id ? "ChevronDown" : "ChevronRight"} style={{ fontSize: 12 }} />
                                            <Icon iconName="Group" style={{ color: '#8764b8', fontSize: 16 }} />
                                            <div style={{ flex: 1 }}>
                                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                                    <div style={{ fontWeight: 600 }}>{group.Title}</div>
                                                    {group.UserCount === 0 && (
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
                                                <div style={{ fontSize: 12, color: '#666' }}>
                                                    {group.UserCount === undefined ? 'Click to load members' : `${group.UserCount} members`}
                                                    {group.Description ? ` • ${group.Description}` : ''}
                                                </div>
                                            </div>
                                        </button>

                                        {expandedGroupId === group.Id && groupMembers[group.Id] && (
                                            <div style={{ paddingLeft: 40, paddingTop: 8, paddingBottom: 8 }}>
                                                {groupMembers[group.Id].length === 0 ? (
                                                    <div style={{ fontSize: 12, color: '#666', fontStyle: 'italic' }}>No members</div>
                                                ) : (
                                                    groupMembers[group.Id].map(member => (
                                                        <div key={member.Id} style={{
                                                            padding: '8px 0',
                                                            borderBottom: '1px solid #f3f2f1',
                                                            display: 'flex',
                                                            alignItems: 'center',
                                                            gap: 8
                                                        }}>
                                                            <Icon iconName="Contact" style={{ fontSize: 14, color: '#605e5c' }} />
                                                            <div>
                                                                <div style={{ fontSize: 13 }}>{member.Title}</div>
                                                                <div style={{ fontSize: 11, color: '#666' }}>{member.Email}</div>
                                                            </div>
                                                        </div>
                                                    ))
                                                )}
                                            </div>
                                        )}
                                    </div>
                                ))}
                        </div>
                    </Panel>

                    <Panel
                        isOpen={isStoragePanelOpen}
                        onDismiss={() => setIsStoragePanelOpen(false)}
                        isLightDismiss={true}
                        type={PanelType.medium}
                        headerText="Storage Metrics"
                        closeButtonAriaLabel="Close"
                    >
                        {siteUsage && (
                            <div style={{ marginBottom: 20, padding: '15px', background: '#f8f9fa', borderRadius: '4px' }}>
                                {siteUsage.storageQuota > 0 ? (
                                    <>
                                        <div style={{ fontSize: 24, fontWeight: 600, color: '#0078d4' }}>
                                            {formatBytes(siteUsage.storageQuota - siteUsage.storageUsed, 2, props.storageFormat)} free
                                        </div>
                                        <div style={{ color: '#666', marginBottom: 10 }}>
                                            of {formatBytes(siteUsage.storageQuota, 2, props.storageFormat)} total quota
                                        </div>
                                        <div style={{ height: 8, background: '#e1dfdd', borderRadius: 4, overflow: 'hidden' }}>
                                            <div style={{
                                                height: '100%',
                                                width: `${(siteUsage.usagePercentage * 100).toFixed(1)}%`,
                                                background: siteUsage.usagePercentage > 0.9 ? '#d13438' : '#0078d4'
                                            }} />
                                        </div>
                                        <div style={{ textAlign: 'right', fontSize: 11, marginTop: 4, color: '#666' }}>
                                            {(siteUsage.usagePercentage * 100).toFixed(2)}% used
                                        </div>
                                    </>
                                ) : (
                                    <>
                                        <div style={{ fontSize: 24, fontWeight: 600, color: '#0078d4' }}>
                                            {formatBytes(siteUsage.storageUsed, 2, props.storageFormat)} used
                                        </div>
                                        <div style={{ color: '#666', marginBottom: 10 }}>
                                            (Total Quota Unknown)
                                        </div>
                                    </>
                                )}
                            </div>
                        )}

                        <p>Storage breakdown by list/library (top items):</p>

                        <table className={styles.permissionTable} style={{ marginTop: 20 }}>
                            <thead>
                                <tr>
                                    <th style={{ width: '40%' }}>Name</th>
                                    <th style={{ width: '20%' }}>Total Size</th>
                                    <th style={{ width: '20%' }}>% of Site</th>
                                    <th style={{ width: '20%' }}>Last Modified</th>
                                </tr>
                            </thead>
                            <tbody>
                                {[...lists]
                                    .sort((a, b) => (b.TotalSize || 0) - (a.TotalSize || 0))
                                    .map(list => {
                                        const size = list.TotalSize || 0;
                                        const percent = siteUsage && siteUsage.storageUsed > 0 ? (size / siteUsage.storageUsed) * 100 : 0;
                                        return (
                                            <tr key={list.Id}>
                                                <td>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                                        <Icon iconName={list.ItemType === 'Library' ? 'DocLibrary' : 'List'} style={{ color: '#0078d4' }} />
                                                        <div>
                                                            <div style={{ fontWeight: 600 }}>{list.Title}</div>
                                                            <div style={{ fontSize: 11, color: '#666' }}>{list.ItemCount} items</div>
                                                        </div>
                                                    </div>
                                                </td>
                                                <td>{formatBytes(size, 2, props.storageFormat)}</td>
                                                <td>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                                        <div style={{ flex: 1, height: 4, background: '#f3f2f1', borderRadius: 2, maxWidth: 50 }}>
                                                            <div style={{ height: '100%', width: `${Math.min(percent, 100)}%`, background: '#0078d4' }} />
                                                        </div>
                                                        <span style={{ fontSize: 11 }}>{percent < 0.1 && percent > 0 ? '<0.1' : percent.toFixed(1)}%</span>
                                                    </div>
                                                </td>
                                                <td style={{ fontSize: 12 }}>
                                                    {list.LastItemModifiedDate ? new Date(list.LastItemModifiedDate).toLocaleDateString() : '-'}
                                                </td>
                                            </tr>
                                        );
                                    })}
                            </tbody>
                        </table>
                    </Panel>
                </div >
            </div >
        </div >
    );
};

export default PermissionViewer;
