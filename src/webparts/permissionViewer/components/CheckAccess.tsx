import * as React from 'react';
import { SearchBox, PrimaryButton, Spinner, SpinnerSize, Persona, PersonaSize, PersonaPresence, Stack, List } from '@fluentui/react';
import { IPermissionService } from '../services/IPermissionService';
import { IUser, IRoleAssignment, IListInfo, IGroup, IItemPermission } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';

export interface ICheckAccessProps {
    permissionService: IPermissionService;
    sitePermissions: IRoleAssignment[]; // Pre-fetched site perms
    lists: IListInfo[]; // Available lists to scan
    contentFontSize?: string;
}

interface IUserReport {
    user: IUser;
    directSitePermissions: string[];
    groupSitePermissions: { group: string; role: string }[];
    listPermissions: { listTitle: string; role: string; type: string }[]; // Type: Direct or Group (if we scan deep)
}

// Helper for robust user matching
const isUserMatch = (user1: IUser, user2: IUser): boolean => {
    if (!user1 || !user2) return false;
    // 1. ID Match (if valid positive ID)
    if (user1.Id > 0 && user2.Id > 0 && user1.Id === user2.Id) return true;
    // 2. LoginName Match
    if (user1.LoginName && user2.LoginName &&
        user1.LoginName.toLowerCase() === user2.LoginName.toLowerCase()) return true;
    // 3. Email Match
    if (user1.Email && user2.Email &&
        user1.Email.toLowerCase() === user2.Email.toLowerCase()) return true;
    return false;
};

const processListPermissions = (
    list: IListInfo,
    perms: IRoleAssignment[],
    userGroupIds: Set<number>,
    selectedUser: IUser,
    results: { listName: string, role: string, type: string }[]
) => {
    if (!selectedUser) return;
    perms.forEach(p => {
        const isDirect = isUserMatch(p.Member, selectedUser);
        const isGroup = userGroupIds.has(p.Member.Id);

        if (isDirect || isGroup) {
            p.RoleDefinitionBindings.forEach(r => {
                results.push({
                    listName: list.Title,
                    role: r.Name,
                    type: isDirect ? 'Direct' : `Group (${p.Member.Title})`
                });
            });
        }
    });
};

const processItemPermissions = (
    list: IListInfo,
    items: IItemPermission[],
    userGroupIds: Set<number>,
    selectedUser: IUser,
    results: { listName: string, role: string, type: string }[]
) => {
    if (!selectedUser) return;

    // Use flatMap/filter chain to reduce nesting depth
    items.forEach(item => {
        item.RoleAssignments.forEach(ra => {
            const isDirect = isUserMatch(ra.Member, selectedUser);
            const isGroup = userGroupIds.has(ra.Member.Id);

            if (isDirect || isGroup) {
                ra.RoleDefinitionBindings.forEach(r => {
                    results.push({
                        listName: `${list.Title} > ${item.Title}`,
                        role: r.Name,
                        type: isDirect ? 'Direct (Item)' : `Group (Item) (${ra.Member.Title})`
                    });
                });
            }
        });
    });
};

export const CheckAccess: React.FunctionComponent<ICheckAccessProps> = (props) => {
    const { permissionService, sitePermissions = [], lists = [], contentFontSize } = props;

    const [searchQuery, setSearchQuery] = React.useState('');
    const [searchResults, setSearchResults] = React.useState<IUser[]>([]);
    const [selectedUser, setSelectedUser] = React.useState<IUser | null>(null);
    const [isSearching, setIsSearching] = React.useState(false);
    const [isScanning, setIsScanning] = React.useState(false);
    const [scanResults, setScanResults] = React.useState<{ listName: string, role: string, type: string }[]>([]);
    // State for user groups
    const [userGroups, setUserGroups] = React.useState<IGroup[]>([]);
    const [hasScanned, setHasScanned] = React.useState(false);

    if (!permissionService) {
        return <div style={{ padding: 20 }}>Service not initialized. Please refresh.</div>;
    }

    const onSearch = async (newValue: string) => {
        setSearchQuery(newValue);
        if (!newValue) {
            setSearchResults([]);
            setSelectedUser(null);
            setScanResults([]);
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

    // Helper for robust user matching (Moved outside)

    const selectUser = async (user: IUser) => {
        setSelectedUser(user);
        setSearchResults([]);
        setSearchQuery(user.Title);
        setScanResults([]);
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

    // Computed Site Access (Direct + Group)
    const getSiteAccess = () => {
        if (!selectedUser) return { direct: [], groups: [] };

        const direct: string[] = [];
        const groups: { group: string; role: string }[] = [];
        const userGroupIds = new Set(userGroups.map(g => g.Id));

        if (Array.isArray(sitePermissions)) {
            sitePermissions.forEach(ra => {
                // Direct Check
                if (ra.Member.PrincipalType === 1 && isUserMatch(ra.Member, selectedUser)) {
                    ra.RoleDefinitionBindings.forEach(r => direct.push(r.Name));
                }
                // Group Check
                if (userGroupIds.has(ra.Member.Id)) {
                    ra.RoleDefinitionBindings.forEach(r => groups.push({ group: ra.Member.Title, role: r.Name }));
                }
            });
        }
        return { direct, groups };
    };

    // Process List and Item functions moved outside for cleanliness

    const runScan = async () => {
        if (!selectedUser) return;
        setIsScanning(true);
        // Clear previous results
        setScanResults([]);

        const results: { listName: string, role: string, type: string }[] = [];
        const uniqueLists = Array.isArray(lists) ? lists.filter(l => l.HasUniqueRoleAssignments) : [];
        const userGroupIds = new Set(userGroups.map(g => g.Id));

        for (const list of uniqueLists) {
            try {
                // 1. Check List Permissions
                const perms = await permissionService.getListRoleAssignments(list.Id, list.Title);
                processListPermissions(list, perms, userGroupIds, selectedUser, results);

                // 2. Check Item Permissions (Deep Scan)
                const items = await permissionService.getUniquePermissionItems(list.Id);
                processItemPermissions(list, items, userGroupIds, selectedUser, results);

            } catch (e) {
                console.error(`Error scanning list ${list.Title}`, e);
            }
        }
        setScanResults(results);
        setIsScanning(false);
        setHasScanned(true);
    };

    const siteAccess = getSiteAccess();

    // Use empty object if styles is undefined
    const safeStyles = styles || {};

    const renderUserCell = (item?: IUser): JSX.Element => (
        <UserResultItem
            item={item}
            onClick={() => item && selectUser(item)}
            className={safeStyles.userResultItem}
        />
    );

    return (
        <div className={safeStyles.checkAccessContainer || ''} style={{ padding: '20px', fontSize: contentFontSize }}>
            <Stack tokens={{ childrenGap: 20 }}>
                {/* Search Header */}
                <Stack>
                    <SearchBox
                        placeholder="Search for a user by name or email..."
                        onChange={(_, val) => onSearch(val || '')}
                        value={searchQuery}
                    />
                    {isSearching && <Spinner size={SpinnerSize.small} label="Searching..." />}

                    {/* Search Results */}
                    {searchResults.length > 0 && (
                        <div style={{ border: '1px solid #edebe9', maxHeight: '200px', overflowY: 'auto' }}>
                            <List
                                items={searchResults}
                                onRenderCell={renderUserCell}
                            />
                        </div>
                    )}
                </Stack>

                {/* Selected User Report */}
                {selectedUser && (
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { border: '1px solid #e1dfdd', padding: '20px', borderRadius: '4px' } }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                            <Persona
                                text={selectedUser.Title}
                                secondaryText={selectedUser.Email}
                                size={PersonaSize.size72}
                                presence={PersonaPresence.none}
                            />
                            <div style={{ flexGrow: 1 }}>
                                <div style={{ fontWeight: 600, fontSize: '16px' }}>{selectedUser.Title}</div>
                                <div style={{ color: '#605e5c' }}>{selectedUser.Email}</div>
                                <div style={{ color: '#605e5c', fontSize: '12px' }}>ID: {selectedUser.Id}</div>
                            </div>
                        </Stack>

                        <div style={{ height: '1px', background: '#edebe9' }} />

                        {/* Site Level Access */}
                        <div>
                            <h3 style={{ marginTop: 0 }}>Site Access</h3>
                            {siteAccess.direct.length === 0 && siteAccess.groups.length === 0 ? (
                                <div>No direct permissions found on the root site.</div>
                            ) : (
                                <ul>
                                    {siteAccess.direct.map((r) => (
                                        <li key={`d-${r}`}><strong>{r}</strong> (Direct)</li>
                                    ))}
                                    {siteAccess.groups.map((g) => (
                                        <li key={`g-${g.group}-${g.role}`}><strong>{g.role}</strong> (via {g.group})</li>
                                    ))}
                                </ul>
                            )}
                        </div>

                        <div style={{ height: '1px', background: '#edebe9' }} />

                        {/* List Level Scan */}
                        <div>
                            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                                <h3 style={{ marginTop: 0, marginBottom: 0 }}>List & Library Access</h3>
                                <PrimaryButton
                                    text={isScanning ? "Scanning..." : "Deep Scan"}
                                    onClick={runScan}
                                    disabled={isScanning}
                                    iconProps={{ iconName: 'Search' }}
                                />
                            </Stack>

                            {isScanning && <div style={{ marginTop: 15 }}><Spinner label="Scanning lists with unique permissions..." /></div>}

                            {!isScanning && scanResults.length > 0 && (
                                <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '15px' }}>
                                    <thead>
                                        <tr style={{ textAlign: 'left', background: '#f3f2f1' }}>
                                            <th style={{ padding: '8px' }}>List Name</th>
                                            <th style={{ padding: '8px' }}>Permission</th>
                                            <th style={{ padding: '8px' }}>Source</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {scanResults.map((res, idx) => (
                                            <tr key={`${res.listName}-${res.role}-${idx}`} style={{ borderBottom: '1px solid #edebe9' }}>
                                                <td style={{ padding: '8px' }}>{res.listName}</td>
                                                <td style={{ padding: '8px' }}>{res.role}</td>
                                                <td style={{ padding: '8px' }}>{res.type}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            )}
                            {!isScanning && !hasScanned && (
                                <div style={{ marginTop: 10, fontStyle: 'italic' }}>
                                    Click "Deep Scan" to check all lists and items with unique permissions.
                                </div>
                            )}
                            {!isScanning && hasScanned && scanResults.length === 0 && (
                                <div style={{ marginTop: 10, color: '#107c10', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <i className="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>
                                    <span>No unique permissions found. This user only has the Site Access shown above.</span>
                                </div>
                            )}
                        </div>

                    </Stack>
                )}
            </Stack>
        </div>
    );
};

// Extracted Component to enable proper interactivity and avoid nesting
const UserResultItem: React.FunctionComponent<{
    item?: IUser,
    onClick: () => void,
    className?: string
}> = ({ item, onClick, className }) => {
    if (!item) return null;



    return (
        <button
            type="button"
            className={className || ''}
            onClick={onClick}
            style={{
                padding: '10px',
                cursor: 'pointer',
                borderBottom: '1px solid #f3f2f1',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                width: '100%',
                background: 'none',
                border: 'none',
                textAlign: 'left'
            }}
        >
            <Persona text={item.Title} secondaryText={item.Email} size={PersonaSize.size24} />
        </button>
    );
};
