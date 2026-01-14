import * as React from 'react';
import { IRoleAssignment } from '../models/IPermissionData';
import { PermissionBadge } from './PermissionBadge';
import { UserPersona } from './UserPersona';
import styles from './PermissionViewer.module.scss';

export interface IListPermissionsTableProps {
    permissions: IRoleAssignment[];
    contentFontSize?: string;
}

export const ListPermissionsTable: React.FunctionComponent<IListPermissionsTableProps> = (props) => {
    const { permissions, contentFontSize } = props;

    // Use default if not provided
    const fontSize = contentFontSize || '14px';

    return (
        <table className={styles.permissionTable} style={{ width: '100%', borderCollapse: 'collapse', marginTop: '16px' }}>
            <thead>
                <tr style={{ background: '#faf9f8', textAlign: 'left' }}>
                    <th style={{ padding: '12px 16px', fontSize: fontSize, fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>User/Group</th>
                    <th style={{ padding: '12px 16px', fontSize: fontSize, fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>Type</th>
                    <th style={{ padding: '12px 16px', fontSize: fontSize, fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>Permission Level</th>
                </tr>
            </thead>
            <tbody>
                {permissions.map((p) => {
                    let userType = 'Security Group';
                    if (p.Member.PrincipalType === 1) userType = 'User';
                    else if (p.Member.PrincipalType === 8) userType = 'SharePoint Group';

                    return (
                        <tr key={p.Member.Id} style={{ borderBottom: '1px solid #f3f2f1' }}>
                            <td style={{ padding: '12px 16px', fontSize: fontSize }}>
                                <UserPersona user={p.Member} fontSize={fontSize} />
                            </td>
                            <td style={{ padding: '12px 16px', fontSize: fontSize }}>
                                {userType}
                            </td>
                            <td style={{ padding: '12px 16px' }}>
                                {p.RoleDefinitionBindings.map(r => (
                                    <PermissionBadge key={r.Id} permission={r.Name} fontSize={fontSize ? `calc(${fontSize} - 2px)` : undefined} />
                                ))}
                            </td>
                        </tr>
                    );
                })}
            </tbody>
        </table>
    );
};
