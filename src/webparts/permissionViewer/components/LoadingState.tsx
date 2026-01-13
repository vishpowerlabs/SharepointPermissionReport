import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './PermissionViewer.module.scss';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';

export interface ILoadingStateProps {
    message: string;
    progress?: number; // 0 to 1
}

export const LoadingState: React.FunctionComponent<ILoadingStateProps> = (props) => {
    return (
        <div className={styles.loadingState} style={{ textAlign: 'center', padding: '60px 32px' }}>
            <Spinner size={SpinnerSize.large} styles={{ circle: { borderColor: '#0078d4' } }} />
            <div className={styles.loadingText} style={{ fontSize: '16px', color: '#605e5c', margin: '12px 0' }}>{props.message}</div>
            <div style={{ width: '300px', margin: '20px auto' }}>
                <ProgressIndicator percentComplete={props.progress} barHeight={4} />
            </div>
            <div style={{ fontSize: '14px', color: '#605e5c', marginTop: '16px' }}>
                This may take a moment for sites with many lists...
            </div>
        </div>
    );
};
