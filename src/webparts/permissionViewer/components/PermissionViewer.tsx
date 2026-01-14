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
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import styles from './PermissionViewer.module.scss';

export interface IPermissionViewerProps {
    spHttpClient: SPHttpClient; // restored
    webUrl: string;
    themeVariant: IReadonlyTheme | undefined;
    headerOpacity?: number;
    showStats?: boolean;
    excludedLists?: string[];
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



    // ... inside component ...

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
            setConfirmScanList(null); // Close confirm dialog only after success
            setDeepScanItems(items);
            setDeepScanListTitle(listTitle);
            setIsDeepScanOpen(true);
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

    return (
        <div className={styles.permissionViewer}>
            <div className={styles.webpartContainer}>
                <Header
                    onRefresh={handleRefresh}
                    isLoading={isLoading}
                    themeVariant={props.themeVariant}
                    opacity={props.headerOpacity ?? 100}
                />

                {/* Default to true if undefined */}
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
                    />
                </div>

                <div className={styles.content}>
                    {isLoading && <LoadingState message={loadingMessage} />}

                    {!isLoading && activeTab === 'site' && (
                        <SitePermissions
                            permissions={filteredSitePermissions}
                            permissionService={permissionService}
                        />
                    )}

                    {!isLoading && activeTab === 'lists' && (
                        <ListPermissions
                            lists={filteredLists}
                            getListPermissions={handleGetListPermissions}
                            onScanItems={handleDeepScan}
                            themeVariant={props.themeVariant}
                        />
                    )}
                </div>
            </div>

            {/* Deep Scan Results Modal */}
            <DeepScanDialog
                isOpen={isDeepScanOpen}
                onDismiss={() => setIsDeepScanOpen(false)}
                listTitle={deepScanListTitle}
                items={deepScanItems}
                onDownload={downloadDeepScanResults}
            />

            <Dialog
                hidden={!confirmScanList}
                onDismiss={() => { if (!isScanning) setConfirmScanList(null); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: isScanning ? 'Deep Scan in Progress...' : 'Start Deep Scan?',
                    subText: isScanning
                        ? `Scanning "${confirmScanList?.title}". This may take a few moments depending on the number of items...`
                        : `This will verify every single item in "${confirmScanList?.title}" to find unique permissions. This might take a while for large lists. Continue?`
                }}
            >
                {isScanning ? (
                    <Spinner size={SpinnerSize.large} label="Scanning for unique permissions..." />
                ) : (
                    <DialogFooter>
                        <PrimaryButton onClick={executeDeepScan} text="Start Scan" />
                        <DefaultButton onClick={() => setConfirmScanList(null)} text="Cancel" />
                    </DialogFooter>
                )}
            </Dialog>
        </div>
    );
};

export default PermissionViewer;
