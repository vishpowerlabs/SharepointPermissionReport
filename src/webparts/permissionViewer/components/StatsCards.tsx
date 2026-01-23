import * as React from 'react';
import { ISiteStats } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IStatsCardsProps {
    stats: ISiteStats;
    highlight?: boolean;
}

export const StatsCards: React.FunctionComponent<IStatsCardsProps> = (props) => {
    const { stats, highlight } = props;

    const cardStyle = highlight ? { transform: 'scale(1.05)', boxShadow: '0 8px 24px rgba(0,0,0,0.2)', border: '2px solid #0078d4' } : {};

    return (
        <div className={styles.statsContainer}>

            <div className={styles.statCard} style={cardStyle}>
                <div className={styles.statIcon} style={{ color: '#8764b8' }}>
                    <Icon iconName="People" />
                </div>
                <div className={styles.statNumber}>{stats.totalGroups}</div>
                <div className={styles.statLabel}>Groups</div>
            </div>
            <div className={styles.statCard} style={cardStyle}>
                <div className={styles.statIcon} style={{ color: '#ca5010' }}>
                    <Icon iconName="Lock" />
                </div>
                <div className={styles.statNumber}>{stats.uniquePermissionsCount}</div>
                <div className={styles.statLabel}>Unique Permissions</div>
            </div>
        </div>
    );
};
