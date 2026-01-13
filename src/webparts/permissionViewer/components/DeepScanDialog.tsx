import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { IItemPermission } from '../models/IPermissionData';
import { PermissionBadge } from './PermissionBadge';
import { UserPersona } from './UserPersona';
import styles from './PermissionViewer.module.scss';

export interface IDeepScanDialogProps {
    isOpen: boolean;
    onDismiss: () => void;
    listTitle: string;
    items: IItemPermission[];
    onDownload: () => void;
}

const renderTypeColumn = (item: IItemPermission) => (
    <div style={{ fontSize: '20px', textAlign: 'center' }}>
        <Icon iconName={item.FileSystemObjectType === 1 ? 'FabricFolder' : 'Page'} style={{ color: item.FileSystemObjectType === 1 ? '#e1c52f' : '#8764b8' }} />
    </div>
);

const renderNameColumn = (item: IItemPermission) => <span title={item.Title}>{item.Title}</span>;

const renderPathColumn = (item: IItemPermission) => <span title={item.ServerRelativeUrl}>{item.ServerRelativeUrl}</span>;

const renderUserColumn = (item: IItemPermission) => (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
        {item.RoleAssignments.map((ra) => (
            <UserPersona key={ra.PrincipalId} user={ra.Member} />
        ))}
    </div>
);

const renderPermissionColumn = (item: IItemPermission) => (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
        {item.RoleAssignments.map((ra) => (
            <div key={ra.PrincipalId} style={{ display: 'flex', gap: '4px' }}>
                {ra.RoleDefinitionBindings.map(r => <PermissionBadge key={r.Id} permission={r.Name} />)}
            </div>
        ))}
    </div>
);

export const DeepScanDialog: React.FunctionComponent<IDeepScanDialogProps> = (props) => {
    const { isOpen, onDismiss, listTitle, items, onDownload } = props;

    const columns: IColumn[] = [
        {
            key: 'type',
            name: 'Type',
            fieldName: 'FileSystemObjectType',
            minWidth: 50,
            maxWidth: 70,
            onRender: renderTypeColumn
        },
        {
            key: 'name',
            name: 'Name',
            fieldName: 'Title',
            minWidth: 150,
            maxWidth: 250,
            isResizable: true,
            onRender: renderNameColumn
        },
        {
            key: 'path',
            name: 'Path',
            fieldName: 'ServerRelativeUrl',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            onRender: renderPathColumn
        },
        {
            key: 'user',
            name: 'User/Group',
            minWidth: 150,
            maxWidth: 200,
            onRender: renderUserColumn
        },
        {
            key: 'roles',
            name: 'Permission',
            minWidth: 150,
            maxWidth: 200,
            onRender: renderPermissionColumn
        }
    ];

    return (
        <Dialog
            hidden={!isOpen}
            onDismiss={onDismiss}
            dialogContentProps={{
                type: DialogType.largeHeader,
                title: `Deep Scan Results: ${listTitle}`,
                subText: `${items.length} items found with unique permissions.`
            }}
            modalProps={{
                isBlocking: false,
                className: styles.deepScanModal,
                styles: { main: { minWidth: '75vw', maxWidth: '75vw' } } // Fallback inline style
            }}
        >
            <div style={{ height: '60vh', overflow: 'auto', border: '1px solid #e1dfdd', borderRadius: '4px' }}>
                <DetailsList
                    items={items}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                />
            </div>
            <DialogFooter>
                <PrimaryButton onClick={onDownload} text="Download CSV" iconProps={{ iconName: 'Download' }} />
                <DefaultButton onClick={onDismiss} text="Close" />
            </DialogFooter>
        </Dialog>
    );
};
