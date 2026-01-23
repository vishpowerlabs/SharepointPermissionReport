import * as React from 'react';
import { Persona, PersonaSize, PersonaPresence, Stack, Text, MessageBar, MessageBarType } from '@fluentui/react';
import { IUser } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';


export interface ISiteAdminsProps {
    users: IUser[];
    isLoading: boolean;
}

export const SiteAdmins: React.FC<ISiteAdminsProps> = ({ users, isLoading }) => {

    const safeStyles = styles || {};

    if (!isLoading && users.length === 0) {
        return (
            <MessageBar messageBarType={MessageBarType.info}>
                No Site Collection Administrators found or you do not have permission to view them.
            </MessageBar>
        );
    }

    const renderRole = (u: IUser) => {
        const roles: string[] = [];
        if (u.IsSiteAdmin) roles.push('Site Admin');
        if (u.IsSiteOwner) roles.push('Site Owner');
        return roles.join(', ') || 'User';
    };

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <div>
                <h3 style={{ marginTop: 0 }}>Site Collection Administrators & Owners</h3>
                <Text>These users have high-level control over the site. Please contact them for access requests.</Text>
            </div>

            <div className={safeStyles.adminGrid || ''}>
                {users.map((user, idx) => (
                    <div key={user.Id || idx} className={safeStyles.adminCard || ''}>
                        <div className={styles.adminHeader}>
                            <Persona
                                text={user.Title}
                                size={PersonaSize.size40}
                                presence={PersonaPresence.none}
                                secondaryText={renderRole(user)}
                            />
                        </div>
                        <div className={styles.adminDetails}>
                            <div style={{ wordBreak: 'break-all' }}>
                                <strong>Email:</strong> {user.Email}
                            </div>
                            <div style={{ wordBreak: 'break-all', fontSize: '11px', color: '#888' }}>
                                <strong>Login:</strong> {user.LoginName}
                            </div>
                        </div>
                    </div>
                ))}
            </div>
        </Stack>
    );
};
