import * as React from 'react';
import { PersonaSize, PersonaCoin } from '@fluentui/react/lib/Persona';
import { IUser } from '../models/IPermissionData';
import styles from './PermissionViewer.module.scss';

export interface IUserPersonaProps {
    user: IUser;
    secondaryText?: string;
    fontSize?: string;
}

export const UserPersona: React.FunctionComponent<IUserPersonaProps> = (props) => {
    const { user, secondaryText, fontSize } = props;

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
                <span className={styles.userName} style={{ fontSize: fontSize || '14px', color: '#323130', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {user.Title}
                    {(user.OrphanStatus === 'Deleted' || (user as any).isDeleted) && (
                        <span style={{
                            backgroundColor: '#fde7e9',
                            color: '#d13438',
                            fontSize: '10px',
                            padding: '2px 6px',
                            borderRadius: '4px',
                            border: '1px solid #d13438',
                            fontWeight: 600,
                            display: 'inline-block'
                        }}>Deleted</span>
                    )}
                    {(user.OrphanStatus === 'Disabled' || (user as any).isDisabled) && (
                        <span style={{
                            backgroundColor: '#f3f2f1',
                            color: '#605e5c',
                            fontSize: '10px',
                            padding: '2px 6px',
                            borderRadius: '4px',
                            border: '1px solid #605e5c',
                            fontWeight: 600,
                            display: 'inline-block'
                        }}>Disabled</span>
                    )}
                </span>
                <span className={styles.userEmail} style={{ fontSize: fontSize ? `calc(${fontSize} - 2px)` : '12px', color: '#605e5c' }}>
                    {description}
                </span>
            </div>
        </div>
    );
};
