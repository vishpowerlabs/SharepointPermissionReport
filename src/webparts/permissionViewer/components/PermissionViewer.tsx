import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PermissionServiceImpl } from '../services/PermissionService';
import { IPermissionService } from '../services/IPermissionService';
import { Header } from './Header';
import { StatsCards } from './StatsCards';
import { SitePermissions } from './SitePermissions';
import { ListPermissions } from './ListPermissions';
import { LoadingState } from './LoadingState';
import { DeepScanDialog } from './DeepScanDialog';
import { IItemPermission, IRoleAssignment, IListInfo, ISiteStats } from '../models/IPermissionData';
import { exportSitePermissions, exportListPermissions, exportDeepScanResults } from '../utils/CsvExport';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react';
import styles from './PermissionViewer.module.scss';

export interface IPermissionViewerProps {
    spHttpClient: SPHttpClient; // restored
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
    const [searchText, setSearchText] = React.useState<string>('');
    const [isExporting, setIsExporting] = React.useState<boolean>(false);
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

    React.useEffect(() => {
        const service = new PermissionServiceImpl(props.spHttpClient, props.webUrl);
        setPermissionService(service);
        loadData(service);
    }, [props.excludedLists]);

    const loadData = async (service: IPermissionService) => {
        setIsLoading(true);
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
            uniquePermissionsCount: uniqueListsCount + (sitePerms.length > 0 ? 1 : 0) // Simplified logic: +1 if site has perms (it always does)
        });

        // Sort lists: Unique first, then Inherited
        listsData.sort((a, b) => {
            if (a.HasUniqueRoleAssignments === b.HasUniqueRoleAssignments) return 0;
            return a.HasUniqueRoleAssignments ? -1 : 1;
        });

        setLists(listsData);
        setFilteredLists(listsData);

        setIsLoading(false);
    };

    const handleRefresh = () => {
        if (permissionService) {
            loadData(permissionService);
        }
    };

    const onSearch = (newValue: string) => {
        setSearchText(newValue);
        const lower = newValue.toLowerCase();

        if (activeTab === 'site') {
            const filtered = sitePermissions.filter(p =>
                p.Member.Title.toLowerCase().includes(lower) ||
                p.Member.Email?.toLowerCase().includes(lower)
            );
            setFilteredSitePermissions(filtered);
        } else {
            const filtered = lists.filter(l =>
                l.Title.toLowerCase().includes(lower) ||
                l.ServerRelativeUrl.toLowerCase().includes(lower)
            );
            setFilteredLists(filtered);
        }
    };

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
                if (searchText) {
                    const lower = searchText.toLowerCase();
                    const filtered = sitePerms.filter(p =>
                        p.Member.Title.toLowerCase().includes(lower) ||
                        p.Member.Email?.toLowerCase().includes(lower)
                    );
                    setFilteredSitePermissions(filtered);
                } else {
                    setFilteredSitePermissions(sitePerms);
                }
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
                        onRefresh={handleRefresh}
                        isLoading={isLoading}
                        themeVariant={props.themeVariant}
                        opacity={props.headerOpacity ?? 100}
                        title={props.webPartTitle}
                        titleFontSize={props.webPartTitleFontSize}
                    />
                )}

                {(props.showStats !== false) && <StatsCards stats={stats} />}

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
                    </Pivot>
                </div>

                <div className={styles.toolbar}>
                    <SearchBox
                        placeholder="Search users or groups..."
                        onChange={(_, newValue) => onSearch(newValue || '')}
                        value={searchText}
                        className={styles.searchBox}
                    />
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
                </div>

                <div className={styles.content}>
                    {isLoading && <LoadingState message={loadingMessage} />}

                    {!isLoading && activeTab === 'site' && (
                        <SitePermissions
                            permissions={filteredSitePermissions}
                            permissionService={permissionService}
                            contentFontSize={props.contentFontSize}
                            onRemovePermission={handleRemoveSitePermission}
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
