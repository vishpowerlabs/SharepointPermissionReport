import * as React from 'react';
import { ISiteStats, ISiteUsage } from '../models/IPermissionData';
import { formatBytes } from '../utils/FormatUtils';
import styles from './PermissionViewer.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';


export interface IStatsCardsProps {
    stats: ISiteStats;
    siteUsage?: ISiteUsage;
    highlight?: boolean;
    onUniquePermissionsClick?: () => void;
    onGroupsClick?: () => void;
    onStorageClick?: () => void;
    storageFormat?: 'Auto' | 'MB' | 'GB' | 'TB';
}

export const StatsCards: React.FunctionComponent<IStatsCardsProps> = (props) => {
    const { stats, siteUsage, highlight, onUniquePermissionsClick, onGroupsClick, onStorageClick, storageFormat } = props;

    const cardStyle = highlight ? { transform: 'scale(1.05)', boxShadow: '0 8px 24px rgba(0,0,0,0.2)', border: '2px solid #0078d4' } : {};

    // Add hover effect style if clickable
    const uniqueCardStyle = {
        ...cardStyle,
        cursor: onUniquePermissionsClick ? 'pointer' : 'default'
    };

    const groupsCardStyle = {
        ...cardStyle,
        cursor: onGroupsClick ? 'pointer' : 'default'
    };

    const storageCardStyle = {
        ...cardStyle,
        cursor: onStorageClick ? 'pointer' : 'default'
    };

    return (
        <div className={styles.statsContainer}>

            <button
                className={styles.statCard}
                style={groupsCardStyle}
                onClick={onGroupsClick}
                disabled={!onGroupsClick}
                type="button"
            >
                <div className={styles.statIcon} style={{ color: '#8764b8' }}>
                    <Icon iconName="People" />
                </div>
                <div className={styles.statNumber}>{stats.totalGroups}</div>
                <div className={styles.statLabel}>Groups</div>
            </button>
            <button
                className={styles.statCard}
                style={uniqueCardStyle}
                onClick={onUniquePermissionsClick}
                disabled={!onUniquePermissionsClick}
                type="button"
            >
                <div className={styles.statIcon} style={{ color: '#ca5010' }}>
                    <Icon iconName="Lock" />
                </div>
                <div className={styles.statNumber}>{stats.uniquePermissionsCount}</div>
                <div className={styles.statLabel}>Unique Permissions</div>
            </button>

            {/* Storage Card */}
            {siteUsage && (
                <button
                    className={styles.statCard}
                    style={storageCardStyle}
                    onClick={onStorageClick}
                    disabled={!onStorageClick}
                    type="button"
                >
                    <div className={styles.statIcon} style={{ color: '#0078d4' }}>
                        <Icon iconName="CloudWeather" />
                    </div>
                    {siteUsage.storageQuota > 0 ? (
                        <>
                            <div className={styles.statNumber} style={{ fontSize: '16px', lineHeight: '1.3' }}>
                                {formatBytes(siteUsage.storageQuota - siteUsage.storageUsed, 2, storageFormat)} free
                            </div>
                            <div className={styles.statLabel} style={{ fontSize: '11px' }}>
                                of {formatBytes(siteUsage.storageQuota, 2, storageFormat)}
                            </div>
                        </>
                    ) : (
                        <>
                            <div className={styles.statNumber} style={{ fontSize: '16px', lineHeight: '1.3' }}>
                                {formatBytes(siteUsage.storageUsed, 2, storageFormat)} used
                            </div>
                            <div className={styles.statLabel} style={{ fontSize: '11px' }}>
                                (Quota Unknown)
                            </div>
                        </>
                    )}
                </button>
            )}

            {/* Activity Card */}
            {siteUsage?.lastItemModifiedDate && (
                <div className={styles.statCard} style={cardStyle}>
                    <div className={styles.statIcon} style={{ color: '#107c10' }}>
                        <Icon iconName="History" />
                    </div>
                    <div className={styles.statNumber} style={{ fontSize: '20px' }}>
                        {new Date(siteUsage.lastItemModifiedDate).toLocaleDateString()}
                    </div>
                    <div className={styles.statLabel}>Last Activity</div>
                </div>
            )}
        </div>
    );
};
