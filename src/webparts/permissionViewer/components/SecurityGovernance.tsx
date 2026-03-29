import * as React from 'react';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './PermissionViewer.module.scss';
import { IUser, ISharingInfo } from '../models/IPermissionData'; // Import ISharingInfo

export interface ISecurityGovernanceProps {
    contentFontSize?: string;
    showExternalUserAudit?: boolean;
    showSharingLinks?: boolean;
    showOrphanedUsers?: boolean;
    onOpenOrphanPanel?: (users: IUser[]) => void;
    // New props for centralized state
    externalUsers?: IUser[];
    sharingLinks?: ISharingInfo[];
    orphanedUsers?: IUser[];
    isLoading?: boolean;
    onRefresh?: () => void;
}

export const SecurityGovernance: React.FunctionComponent<ISecurityGovernanceProps> = (props) => {
    const [activeModal, setActiveModal] = React.useState<string | null>(null);

    const externalUsers = props.externalUsers || [];
    const sharingLinks = props.sharingLinks || [];
    const orphanedUsers = props.orphanedUsers || [];
    const isLoading = props.isLoading || false;

    const getCardData = () => {
        return [
            {
                id: 'audit',
                title: 'External User Audit',
                icon: '🌐',
                count: externalUsers.length,
                label: 'Guest Users found',
                statClass: styles.govStatInfo,
                details: [
                    `• ${externalUsers.length} external accounts`,
                    '• Review regularly'
                ],
                modalTitle: 'External User Audit Details',
                data: externalUsers
            },
            {
                id: 'links',
                title: 'Sharing Links Report',
                icon: '🔗',
                count: sharingLinks.length,
                label: 'Sharing Links',
                statClass: styles.govStatWarning,
                details: [
                    `• ${sharingLinks.length} active links`,
                    '• Check for anonymous access'
                ],
                modalTitle: 'Sharing Links Details',
                data: sharingLinks
            },
            {
                id: 'orphaned',
                title: 'Orphaned Users',
                icon: '⚠️',
                count: orphanedUsers.length,
                label: 'Users detected',
                statClass: styles.govStatAlert,
                details: [
                    `• ${orphanedUsers.length} potentially orphaned`,
                    '• Users not in AD or disabled'
                ],
                modalTitle: 'Orphaned Users Details',
                data: orphanedUsers
            }
        ];
    };

    const cards = getCardData().filter(c => {
        if (c.id === 'audit') return props.showExternalUserAudit !== false;
        if (c.id === 'links') return props.showSharingLinks !== false;
        if (c.id === 'orphaned') return props.showOrphanedUsers !== false;
        return true;
    });

    const activeCard = cards.find(c => c.id === activeModal);

    if (isLoading) {
        return <Spinner size={SpinnerSize.large} label="Loading governance reports..." />;
    }

    return (
        <div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '15px' }}>
                <PrimaryButton
                    text="Refresh Governance Data"
                    iconProps={{ iconName: 'Refresh' }}
                    onClick={props.onRefresh}
                    disabled={props.isLoading}
                />
            </div>
            <div className={styles.governanceGrid}>
                {cards.map((card, i) => (
                    <div key={card.id} className={styles.governanceCard}>
                        <div className={styles.govIcon}>{card.icon}</div>
                        <h3 className={styles.govTitle} style={{ fontSize: props.contentFontSize }}>{card.title}</h3>
                        <div className={`${styles.govStat} ${card.statClass}`}>
                            <span>{card.count}</span> {card.label}
                        </div>
                        <ul className={styles.govList}>
                            {card.details.map((d, idx) => (
                                <li key={`${idx}-${d.substring(0, 5)}`}>{d}</li>
                            ))}
                        </ul>
                        <PrimaryButton
                            text="View Details"
                            className={styles.govBtn}
                            onClick={() => {
                                if (card.id === 'orphaned' && props.onOpenOrphanPanel) {
                                    props.onOpenOrphanPanel(orphanedUsers);
                                } else {
                                    setActiveModal(card.id);
                                }
                            }}
                        />
                    </div>
                ))}

                <Dialog
                    hidden={!activeModal}
                    onDismiss={() => setActiveModal(null)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: activeCard?.modalTitle,
                        styles: { subText: { fontSize: props.contentFontSize } }
                    }}
                    modalProps={{
                        isBlocking: false,
                        styles: { main: { maxWidth: 700 } }
                    }}
                >
                    {activeCard && (
                        <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                            <table className={styles.permissionTable}>
                                <thead>
                                    <tr>
                                        <th>Name</th>
                                        <th>Login Name / Email</th>
                                        <th>Type</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {activeCard.data.length === 0 ? (
                                        <tr>
                                            <td colSpan={3} style={{ textAlign: 'center', padding: '20px' }}>No items found.</td>
                                        </tr>
                                    ) : (
                                        activeCard.data.map((item: any, idx: number) => {
                                            if (activeCard.id === 'links') {
                                                const link = item as ISharingInfo;
                                                return (
                                                    <tr key={`${link.documentUrl}-${link.linkType}-${link.sharedWith.join('')}`}>
                                                        <td>
                                                            <a href={link.documentUrl} target="_blank" data-interception="off" rel="noopener noreferrer">
                                                                {link.documentName}
                                                            </a>
                                                        </td>
                                                        <td>
                                                            {link.sharedWith.join(", ")}
                                                            {link.linkType && <div style={{ fontSize: '11px', color: '#666' }}>({link.linkType})</div>}
                                                        </td>
                                                        <td>Sharing Link</td>
                                                    </tr>
                                                );
                                            } else {
                                                const u = item as IUser;
                                                let typeLabel = 'Orphaned User';
                                                if (activeCard.id === 'audit') typeLabel = 'External User';
                                                else if (activeCard.id === 'orphaned' && u.OrphanStatus) {
                                                    typeLabel = u.OrphanStatus; // 'Deleted' or 'Disabled'
                                                }

                                                // Handling standard user display
                                                return (
                                                    <tr key={u.Id || u.LoginName || idx}>
                                                        <td>{u.Title}</td>
                                                        <td>{u.Email || u.LoginName}</td>
                                                        <td>{typeLabel}</td>
                                                    </tr>
                                                );
                                            }
                                        })
                                    )}
                                </tbody>
                            </table>
                        </div>
                    )}
                    <DialogFooter>
                        <PrimaryButton onClick={() => setActiveModal(null)} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        </div>
    );
};
