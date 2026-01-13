import * as React from 'react';
import styles from './PermissionViewer.module.scss';

export interface IPermissionBadgeProps {
    permission: string;
    isInheritanceStatus?: boolean; // if true, it's Unique/Inherited badge
}

export const PermissionBadge: React.FunctionComponent<IPermissionBadgeProps> = (props) => {
    const { permission, isInheritanceStatus } = props;

    let text = permission;

    if (isInheritanceStatus) {
        if (permission === 'Unique') {
            text = '🔒 UNIQUE';
        } else {
            text = '⬆️ INHERITED';
        }
    }

    const getStyleKey = (perm: string): string => {
        // Normalize to camelCase
        return 'badge' + perm.replaceAll(' ', '');
    };

    const styleName = isInheritanceStatus
        ? (permission === 'Unique' ? 'badgeUnique' : 'badgeInherited')
        : getStyleKey(permission);

    // Use type assertion to avoid TS7053 since we know the keys exist from SCSS
    const styleClass = (styles as any)[styleName] || styles.permissionBadge;

    return (
        <span className={`${styles.permissionBadge} ${styleClass}`}
            style={!styleClass ? getBadgeStyle(permission, isInheritanceStatus) : undefined}>
            {text}
        </span>
    );
};

const getBadgeStyle = (permission: string, isInheritanceStatus?: boolean): React.CSSProperties => {
    const base: React.CSSProperties = {
        padding: '4px 12px',
        borderRadius: '12px',
        fontSize: '12px',
        fontWeight: 600,
        border: '1px solid',
        display: 'inline-block'
    };

    if (isInheritanceStatus) {
        if (permission === 'Unique') return { ...base, background: '#fef6f1', color: '#ca5010', borderColor: '#ca5010' };
        return { ...base, background: '#f3f2f1', color: '#605e5c', borderColor: '#8a8886' };
    }

    if (permission.includes('Full Control')) return { ...base, background: '#fde7e9', color: '#d13438', borderColor: '#d13438' };
    if (permission.includes('Edit')) return { ...base, background: '#fef6f1', color: '#ca5010', borderColor: '#ca5010' };
    if (permission.includes('Contribute')) return { ...base, background: '#e6f2ff', color: '#0078d4', borderColor: '#0078d4' };
    if (permission.includes('Read')) return { ...base, background: '#e6f7e6', color: '#107c10', borderColor: '#107c10' };
    if (permission.includes('Limited Access')) return { ...base, background: '#f3f2f1', color: '#605e5c', borderColor: '#8a8886' };

    return { ...base, background: '#f3f2f1', color: '#605e5c', borderColor: '#605e5c' };
}
