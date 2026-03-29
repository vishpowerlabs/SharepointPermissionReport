
import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { DefaultButton, Stack, Spinner, SpinnerSize, MessageBar, MessageBarType, Link, Icon, TooltipHost } from '@fluentui/react';
import { IOversharedFolder } from '../models/IPermissionData';
import { exportFolderOversharingResults } from '../utils/CsvExport';

export interface IFolderOversharingPanelProps {
    isOpen: boolean;
    onDismiss: () => void;
    listName: string;
    isScanning: boolean;
    results: IOversharedFolder[];
    onCheckRootItems: () => void;
    onScanContents?: (folderUrl: string) => void;
    isScanContentsLoading?: boolean;
}

export const FolderOversharingPanel: React.FunctionComponent<IFolderOversharingPanelProps> = (props) => {
    const { isOpen, onDismiss, listName, isScanning, results, onCheckRootItems, onScanContents, isScanContentsLoading } = props;

    const [scanningFolderUrl, setScanningFolderUrl] = React.useState<string | null>(null);

    const handleScanContents = (folderUrl: string) => {
        setScanningFolderUrl(folderUrl);
        if (onScanContents) {
            onScanContents(folderUrl);
        }
    };

    // Reset local scan state when global loading stops (assuming completion)
    // Actually, parent should handle loading state, but we need to know WHICH folder is scanning to show spinner on button?
    // For simplicity, let's just use a global spinner overlay or assume the user waits.
    // Or better: Use 'scanningFolderUrl' local state to show spinner on that specific row button?
    // The prop 'isScanContentsLoading' passed from parent can signal when to clear 'scanningFolderUrl'.

    React.useEffect(() => {
        if (!isScanContentsLoading) {
            setScanningFolderUrl(null);
        }
    }, [isScanContentsLoading]);


    const columns: IColumn[] = [
        {
            key: 'name', name: 'Folder Name', fieldName: 'Name', minWidth: 150, maxWidth: 200, isResizable: true,
            onRender: (item: IOversharedFolder) => (
                <Link href={item.Path} target="_blank" data-interception="off">{item.Name}</Link>
            )
        },
        {
            key: 'shared', name: 'Shared With', fieldName: 'SharedWith', minWidth: 150, maxWidth: 200, isResizable: true,
            onRender: (item: IOversharedFolder) => (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }}>
                    <TooltipHost content="Everyone except external users">
                        <Icon iconName="Globe" styles={{ root: { color: '#0078d4', fontSize: 16 } }} />
                    </TooltipHost>
                </Stack>
            )
        },
        { key: 'perms', name: 'Permissions', fieldName: 'Permissions', minWidth: 100, maxWidth: 150, isResizable: true },
        {
            key: 'actions', name: 'Actions', minWidth: 120, maxWidth: 150,
            onRender: (item: IOversharedFolder) => {
                const isThisScanning = scanningFolderUrl === item.Path;
                return (
                    <DefaultButton
                        text={isThisScanning ? "Scanning..." : "Scan Contents"}
                        iconProps={isThisScanning ? { iconName: 'Sync', styles: { root: { animation: 'spin 2s linear infinite' } } } : { iconName: 'FabricFolderSearch' }}
                        onClick={() => handleScanContents(item.Path)}
                        disabled={isScanning || isScanContentsLoading}
                        styles={{ root: { height: 32, fontSize: 12 } }}
                    />
                );
            }
        }
    ];

    const onExport = () => {
        exportFolderOversharingResults(results, listName);
    };

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            headerText={`Oversharing Check: ${listName}`}
            closeButtonAriaLabel="Close"
            isBlocking={isScanning || !!isScanContentsLoading}
        >
            <Stack tokens={{ childrenGap: 20 }}>
                <p>
                    Scanning folders for broad sharing ("Everyone", "Anonymous").
                    <br />
                    <b>Note:</b> If a folder is shared, its subfolders are skipped to reduce noise (they inherit the risk).
                    Use <b>"Scan Contents"</b> to dig deeper into a specific folder.
                </p>

                <Stack horizontal tokens={{ childrenGap: 10 }}>
                    <DefaultButton
                        text="Check Root Items (Shallow)"
                        iconProps={{ iconName: 'FabricFolder' }}
                        onClick={onCheckRootItems}
                        disabled={isScanning || !!isScanContentsLoading}
                    />
                </Stack>

                {isScanning && (
                    <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                        <Spinner size={SpinnerSize.large} label="Scanning folders..." />
                    </Stack>
                )}

                {!isScanning && results.length === 0 && (
                    <MessageBar messageBarType={MessageBarType.success}>
                        No folders found shared with "Everyone" or "Anonymous".
                    </MessageBar>
                )}

                {!isScanning && results.length > 0 && (
                    <Stack tokens={{ childrenGap: 10 }}>
                        <MessageBar messageBarType={MessageBarType.warning}>
                            Found {results.length} folders with broad access.
                        </MessageBar>

                        <div style={{ marginBottom: 10 }}>
                            <DefaultButton
                                text="Export to CSV"
                                iconProps={{ iconName: 'Download' }}
                                onClick={onExport}
                            />
                        </div>

                        <DetailsList
                            items={results}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                        />
                    </Stack>
                )}
            </Stack>
            <style>{`
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            `}</style>
        </Panel>
    );
};
