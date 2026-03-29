import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './PermissionViewer.module.scss';
import { ProgressIndicator, Stack, Text, FontWeights } from '@fluentui/react';

export interface ILoadingStateProps {
    message: string;
    progress?: number; // 0 to 1
}

export const LoadingState: React.FunctionComponent<ILoadingStateProps> = (props) => {
    return (
        <Stack
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 20 }}
            styles={{
                root: {
                    padding: '80px 40px',
                    minHeight: '300px',
                    backgroundColor: '#faf9f8', // NeutralLighter
                    borderRadius: '8px'
                }
            }}
        >
            <Spinner size={SpinnerSize.large} label={props.message} ariaLive="assertive" labelPosition="top" />

            {props.progress !== undefined && (
                <Stack styles={{ root: { width: '300px', maxWidth: '100%' } }}>
                    <ProgressIndicator percentComplete={props.progress} barHeight={4} />
                    <Text variant="small" styles={{ root: { marginTop: 8, textAlign: 'center', color: '#605e5c' } }}>
                        {(props.progress * 100).toFixed(0)}% Complete
                    </Text>
                </Stack>
            )}

            <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic' } }}>
                Use the "Stop Scan" button if this takes too long.
            </Text>
        </Stack>
    );
};
