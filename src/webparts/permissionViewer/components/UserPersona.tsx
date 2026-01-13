import * as React from 'react';
import { PersonaSize, PersonaCoin } from '@fluentui/react/lib/Persona';
import { IUser } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';

export interface IUserPersonaProps {
    user: IUser;
    secondaryText?: string;
}

export const UserPersona: React.FunctionComponent<IUserPersonaProps> = (props) => {
    const { user, secondaryText } = props;

    // Principal Type: 1=User, 4=Security Group, 8=SharePoint Group
    const isGroup = user.PrincipalType === 8 || user.PrincipalType === 4;

    let description = secondaryText;
    if (!description) {
        if (user.Email) {
            description = user.Email;
        } else if (isGroup) {
            description = 'Group';
        } else {
            description = '';
        }
    }

    return (
        <div className={styles.userCell} style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <PersonaCoin
                text={user.Title}
                size={PersonaSize.size32}
                coinSize={32}
            // Using explicit styles here or relying on className if feasible, but coinInner isn't a valid prop in some versions
            />
            <div className={styles.userInfo} style={{ display: 'flex', flexDirection: 'column' }}>
                <span className={styles.userName} style={{ fontSize: '14px', color: '#323130', fontWeight: 600 }}>{user.Title}</span>
                <span className={styles.userEmail} style={{ fontSize: '12px', color: '#605e5c' }}>
                    {description}
                </span>
            </div>
        </div>
    );
};
