import * as React from 'react';
import { ISiteStats } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IStatsCardsProps {
    stats: ISiteStats;
}

export const StatsCards: React.FunctionComponent<IStatsCardsProps> = (props) => {
    const { stats } = props;

    return (
        <div className={styles.statsContainer}>
            <div className={styles.statCard}>
                <div className={styles.statIcon} style={{ color: '#0078d4' }}>
                    <Icon iconName="Group" />
                </div>
                <div className={styles.statNumber}>{stats.totalUsers}</div>
                <div className={styles.statLabel}>Total Users</div>
            </div>
            <div className={styles.statCard}>
                <div className={styles.statIcon} style={{ color: '#8764b8' }}>
                    <Icon iconName="People" />
                </div>
                <div className={styles.statNumber}>{stats.totalGroups}</div>
                <div className={styles.statLabel}>Groups</div>
            </div>
            <div className={styles.statCard}>
                <div className={styles.statIcon} style={{ color: '#ca5010' }}>
                    <Icon iconName="Lock" />
                </div>
                <div className={styles.statNumber}>{stats.uniquePermissionsCount}</div>
                <div className={styles.statLabel}>Unique Permissions</div>
            </div>
        </div>
    );
};
