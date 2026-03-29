import * as React from 'react';
import { SearchBox, PrimaryButton, Spinner, SpinnerSize, Persona, PersonaSize, PersonaPresence, Stack, List } from '@fluentui/react';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode, IGroup } from '@fluentui/react/lib/DetailsList';
import { IPermissionService } from '../services/IPermissionService';
import { IUser, IRoleAssignment, IListInfo, IGroup as ISharePointGroup, IItemPermission } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';
import { downloadCsv } from '../utils/CsvExport';

export interface ICheckAccessProps {
    permissionService: IPermissionService;
    sitePermissions: IRoleAssignment[]; // Pre-fetched site perms
    lists: IListInfo[]; // Available lists to scan
    contentFontSize?: string;
}

interface IAccessResult {
    key: string;
    scope: string; // Site, List, Item
    listName: string; // Grouping Key
    location: string; // Item Name / URL
    role: string;
    source: string; // Direct vs Group
    url?: string;
}

// Helper for robust user matching
const isUserMatch = (user1: IUser, user2: IUser): boolean => {
    if (!user1 || !user2) return false;
    if (user1.Id > 0 && user2.Id > 0 && user1.Id === user2.Id) return true;
    if (user1.LoginName && user2.LoginName &&
        user1.LoginName.toLowerCase() === user2.LoginName.toLowerCase()) return true;
    if (user1.Email && user2.Email &&
        user1.Email.toLowerCase() === user2.Email.toLowerCase()) return true;
    return false;
};

export const CheckAccess: React.FunctionComponent<ICheckAccessProps> = (props) => {
    const { permissionService, sitePermissions = [], lists = [], contentFontSize } = props;

    const [searchQuery, setSearchQuery] = React.useState('');
    const [searchResults, setSearchResults] = React.useState<IUser[]>([]);
    const [selectedUser, setSelectedUser] = React.useState<IUser | null>(null);
    const [isSearching, setIsSearching] = React.useState(false);

    // Scan State
    const [isScanning, setIsScanning] = React.useState(false);
    const [scanResults, setScanResults] = React.useState<IAccessResult[]>([]);
    const [resultGroups, setResultGroups] = React.useState<IGroup[]>([]);
    const [hasScanned, setHasScanned] = React.useState(false);

    // State for user groups
    const [userGroups, setUserGroups] = React.useState<ISharePointGroup[]>([]);

    if (!permissionService) {
        return <div style={{ padding: 20 }}>Service not initialized. Please refresh.</div>;
    }

    const onSearch = async (newValue: string) => {
        setSearchQuery(newValue);
        if (!newValue) {
            setSearchResults([]);
            setSelectedUser(null);
            setScanResults([]);
            setResultGroups([]);
            setHasScanned(false);
            return;
        }
        if (newValue.length > 2) {
            setIsSearching(true);
            try {
                const results = await permissionService.searchUsers(newValue);
                setSearchResults(results);
            } catch (err) {
                console.error("Error searching users", err);
            } finally {
                setIsSearching(false);
            }
        } else {
            setSearchResults([]);
        }
    };

    const selectUser = async (user: IUser) => {
        setSelectedUser(user);
        setSearchResults([]);
        setSearchQuery(user.Title);
        setScanResults([]);
        setResultGroups([]);
        setHasScanned(false);

        // Fetch groups for selected user
        if (user.LoginName) {
            try {
                const groups = await permissionService.getUserGroups(user.LoginName);
                setUserGroups(groups);
            } catch (e) {
                console.error("Error fetching user groups", e);
                setUserGroups([]);
            }
        } else {
            setUserGroups([]);
        }
    };

    const runFullReport = async () => {
        if (!selectedUser) return;
        setIsScanning(true);
        setScanResults([]);
        setResultGroups([]);

        const results: IAccessResult[] = [];
        const userGroupIds = new Set(userGroups.map(g => g.Id));

        // 1. Site Access (Root)
        sitePermissions.forEach(ra => {
            const isDirect = ra.Member.PrincipalType === 1 && isUserMatch(ra.Member, selectedUser);
            const isGroup = userGroupIds.has(ra.Member.Id);

            if (isDirect || isGroup) {
                ra.RoleDefinitionBindings.forEach(r => {
                    results.push({
                        key: `site-${ra.Member.Id}-${r.Id}`,
                        scope: 'Site',
                        listName: 'Site Collection Root',
                        location: '/',
                        role: r.Name,
                        source: isDirect ? 'Direct' : `Group (${ra.Member.Title})`
                    });
                });
            }
        });

        // 2. Lists & Items
        const uniqueLists = Array.isArray(lists) ? lists.filter(l => l.HasUniqueRoleAssignments) : [];

        for (const list of uniqueLists) {
            try {
                // List Permissions
                const perms = await permissionService.getListRoleAssignments(list.Id, list.Title);
                perms.forEach(p => {
                    const isDirect = isUserMatch(p.Member, selectedUser);
                    const isGroup = userGroupIds.has(p.Member.Id);

                    if (isDirect || isGroup) {
                        p.RoleDefinitionBindings.forEach(r => {
                            results.push({
                                key: `list-${list.Id}-${p.Member.Id}-${r.Id}`,
                                scope: 'List/Library',
                                listName: list.Title,
                                location: list.ServerRelativeUrl,
                                role: r.Name,
                                source: isDirect ? 'Direct' : `Group (${p.Member.Title})`,
                                url: list.ServerRelativeUrl
                            });
                        });
                    }
                });

                // Deep Scan Items
                const items = await permissionService.getUniquePermissionItems(list.Id);
                items.forEach(item => {
                    item.RoleAssignments.forEach(ra => {
                        const isDirect = isUserMatch(ra.Member, selectedUser);
                        const isGroup = userGroupIds.has(ra.Member.Id);

                        if (isDirect || isGroup) {
                            ra.RoleDefinitionBindings.forEach(r => {
                                results.push({
                                    key: `item-${item.Id}-${ra.PrincipalId}-${r.Id}`,
                                    scope: item.FileSystemObjectType === 1 ? 'Folder' : 'File',
                                    listName: list.Title,
                                    location: item.ServerRelativeUrl,
                                    role: r.Name,
                                    source: isDirect ? 'Direct' : `Group (${ra.Member.Title})`,
                                    url: item.ServerRelativeUrl
                                });
                            });
                        }
                    });
                });

            } catch (e) {
                console.error(`Error scanning list ${list.Title}`, e);
            }
        }

        // Grouping Logic
        // Sort by ListName first
        results.sort((a, b) => {
            if (a.listName === 'Site Collection Root') return -1;
            if (b.listName === 'Site Collection Root') return 1;
            return a.listName.localeCompare(b.listName);
        });

        // Create Groups
        const groups: IGroup[] = [];
        let currentGroup: IGroup | null = null;

        results.forEach((item, index) => {
            if (!currentGroup || currentGroup.name !== item.listName) {
                if (currentGroup) {
                    // Explicit cast to avoid typescript 'never' inference weirdness
                    const group = currentGroup as IGroup;
                    group.count = index - group.startIndex;
                    groups.push(group);
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
            group.count = results.length - group.startIndex;
            groups.push(group);
        }

        setScanResults(results);
        setResultGroups(groups);
        setIsScanning(false);
        setHasScanned(true);
    };

    const handleExport = () => {
        if (scanResults.length === 0 || !selectedUser) return;

        let csv = "Scope,Grouping,Location,Permission,Source\n";
        scanResults.forEach(r => {
            csv += `"${r.scope}","${r.listName}","${r.location}","${r.role}","${r.source}"\n`;
        });

        const fileName = `AccessReport_${selectedUser.Title.replace(/[^a-z0-9]/gi, '_')}.csv`;
        downloadCsv(csv, fileName);
    };

    const columns: IColumn[] = [
        { key: 'scope', name: 'Scope', fieldName: 'scope', minWidth: 60, maxWidth: 80, isResizable: true },
        { key: 'location', name: 'Location', fieldName: 'location', minWidth: 200, maxWidth: 400, isResizable: true },
        { key: 'role', name: 'Permission', fieldName: 'role', minWidth: 100, maxWidth: 150, isResizable: true },
        { key: 'source', name: 'Source', fieldName: 'source', minWidth: 150, maxWidth: 200, isResizable: true },
    ];

    // Computed Site Access for Summary Header
    const getSiteAccessSummary = () => {
        if (!selectedUser) return [];
        const roles = new Set<string>();
        const userGroupIds = new Set(userGroups.map(g => g.Id));

        sitePermissions.forEach(ra => {
            if ((ra.Member.PrincipalType === 1 && isUserMatch(ra.Member, selectedUser)) || userGroupIds.has(ra.Member.Id)) {
                ra.RoleDefinitionBindings.forEach(r => roles.add(r.Name));
            }
        });
        return Array.from(roles);
    };

    const siteRoles = getSiteAccessSummary();

    const renderUserCell = (item?: IUser): JSX.Element => (
        <UserResultItem
            item={item}
            onClick={() => item && selectUser(item)}
            className={props.contentFontSize ? undefined : styles.userResultItem} // Use undefined if using inline styles mostly, or fix scss later
        />
    );

    return (
        <div style={{ padding: '20px', fontSize: contentFontSize }}>
            <Stack tokens={{ childrenGap: 20 }}>
                {/* Search Header */}
                <Stack>
                    <SearchBox
                        placeholder="Search for a user by name or email..."
                        onChange={(_, val) => onSearch(val || '')}
                        value={searchQuery}
                        styles={{ root: { maxWidth: 400 } }}
                    />
                    {isSearching && <Spinner size={SpinnerSize.small} label="Searching..." style={{ marginTop: 10 }} />}

                    {/* Search Results */}
                    {searchResults.length > 0 && (
                        <div style={{ border: '1px solid #edebe9', maxHeight: '200px', overflowY: 'auto', maxWidth: 400, position: 'absolute', zIndex: 1000, background: 'white', marginTop: 32 }}>
                            <List items={searchResults} onRenderCell={renderUserCell} />
                        </div>
                    )}
                </Stack>

                {/* Selected User Report */}
                {selectedUser && (
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { border: '1px solid #e1dfdd', padding: '20px', borderRadius: '4px', background: 'white' } }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }}>
                            <Persona
                                text={selectedUser.Title}
                                secondaryText={selectedUser.Email}
                                size={PersonaSize.size72}
                                presence={PersonaPresence.none}
                            />
                            <div style={{ flexGrow: 1 }}>
                                <div style={{ fontWeight: 600, fontSize: '18px' }}>{selectedUser.Title}</div>
                                <div style={{ color: '#605e5c' }}>{selectedUser.Email}</div>
                                <div style={{ color: '#605e5c', fontSize: '12px', marginTop: 4 }}>
                                    Site Level Access: {siteRoles.length > 0 ? <strong>{siteRoles.join(', ')}</strong> : <span>None (or limited to specific lists)</span>}
                                </div>
                            </div>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <PrimaryButton
                                    text={isScanning ? "Generating Report..." : "Generate Full Access Report"}
                                    onClick={runFullReport}
                                    disabled={isScanning}
                                    iconProps={{ iconName: 'ReportDocument' }}
                                />
                                {hasScanned && scanResults.length > 0 && (
                                    <PrimaryButton
                                        text="Export Report"
                                        onClick={handleExport}
                                        iconProps={{ iconName: 'Download' }}
                                        styles={{ root: { backgroundColor: '#107c10', borderColor: '#107c10' } }}
                                    />
                                )}
                            </Stack>
                        </Stack>

                        <div style={{ height: '1px', background: '#edebe9' }} />

                        {/* Report Area */}
                        <div>
                            {isScanning && (
                                <div style={{ padding: 40, textAlign: 'center' }}>
                                    <Spinner size={SpinnerSize.large} label="Scanning entire site structure... This may take a minute." />
                                </div>
                            )}

                            {!isScanning && hasScanned && (
                                scanResults.length === 0 ? (
                                    <div style={{ padding: 20, textAlign: 'center', color: '#605e5c' }}>
                                        <div style={{ fontSize: 40, marginBottom: 10 }}>🔍</div>
                                        No access found. This user appears to have no permissions on this site or its lists.
                                    </div>
                                ) : (
                                    <div>
                                        <div style={{ marginBottom: 10, fontWeight: 600 }}>Found {scanResults.length} permission assignments</div>
                                        <DetailsList
                                            items={scanResults}
                                            groups={resultGroups}
                                            columns={columns}
                                            selectionMode={SelectionMode.none}
                                            layoutMode={DetailsListLayoutMode.justified}
                                            groupProps={{
                                                showEmptyGroups: true,
                                            }}
                                            styles={{ root: { overflowX: 'visible' } }} // Ensure tooltips/content don't clip awkwardly
                                        />
                                    </div>
                                )
                            )}

                            {!isScanning && !hasScanned && (
                                <div style={{ padding: 20, textAlign: 'center', color: '#605e5c', fontStyle: 'italic' }}>
                                    Click "Generate Full Access Report" to scan all Site, List, Library, Folder, and Item permissions for this user.
                                </div>
                            )}
                        </div>
                    </Stack>
                )}
            </Stack>
        </div>
    );
};

// Extracted Component
const UserResultItem: React.FunctionComponent<{
    item?: IUser,
    onClick: () => void,
    className?: string
}> = ({ item, onClick, className }) => {
    if (!item) return null;

    return (
        <div
            className={className || ''}
            onClick={onClick}
            style={{
                padding: '10px',
                cursor: 'pointer',
                borderBottom: '1px solid #f3f2f1',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                width: '100%'
            }}
        >
            <Persona text={item.Title} secondaryText={item.Email} size={PersonaSize.size24} />
        </div>
    );
};
