import * as React from 'react';
import { IRoleAssignment } from '../models/IPermissionData';
import { PermissionBadge } from './PermissionBadge';
import { UserPersona } from './UserPersona';
import styles from './PermissionViewer.module.scss';

export interface IListPermissionsTableProps {
    permissions: IRoleAssignment[];
}

export const ListPermissionsTable: React.FunctionComponent<IListPermissionsTableProps> = (props) => {
    const { permissions } = props;

    return (
        <table className={styles.permissionTable} style={{ width: '100%', borderCollapse: 'collapse', marginTop: '16px' }}>
            <thead>
                <tr style={{ background: '#faf9f8', textAlign: 'left' }}>
                    <th style={{ padding: '12px 16px', fontSize: '14px', fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>User/Group</th>
                    <th style={{ padding: '12px 16px', fontSize: '14px', fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>Type</th>
                    <th style={{ padding: '12px 16px', fontSize: '14px', fontWeight: 600, color: '#323130', borderBottom: '1px solid #e1dfdd' }}>Permission Level</th>
                </tr>
            </thead>
            <tbody>
                {permissions.map((p) => {
                    let userType = 'Security Group';
                    if (p.Member.PrincipalType === 1) userType = 'User';
                    else if (p.Member.PrincipalType === 8) userType = 'SharePoint Group';

                    return (
                        <tr key={p.Member.Id} style={{ borderBottom: '1px solid #f3f2f1' }}>
                            <td style={{ padding: '12px 16px' }}>
                                <UserPersona user={p.Member} />
                            </td>
                            <td style={{ padding: '12px 16px' }}>
                                {userType}
                            </td>
                            <td style={{ padding: '12px 16px' }}>
                                {p.RoleDefinitionBindings.map(r => <PermissionBadge key={r.Id} permission={r.Name} />)}
                            </td>
                        </tr>
                    );
                })}
            </tbody>
        </table>
    );
};
