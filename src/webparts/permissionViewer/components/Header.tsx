import * as React from 'react';
import styles from './PermissionViewer.module.scss';
import { IReadonlyTheme } from '@microsoft/sp-component-base';


export interface IHeaderProps {
    themeVariant: IReadonlyTheme | undefined;
    opacity?: number;
    title?: string;
    titleFontSize?: string;
}

const hexToRgba = (hex: string, alpha: number) => {
    // Basic validation
    if (!hex || !/^#([A-Fa-f0-9]{3}){1,2}$/.test(hex)) {
        return `rgba(0, 120, 212, ${alpha})`; // Fallback to blue
    }

    let c = hex.substring(1).split('');
    if (c.length === 3) {
        c = [c[0], c[0], c[1], c[1], c[2], c[2]];
    }
    const val = Number.parseInt(c.join(''), 16);
    return `rgba(${(val >> 16) & 255}, ${(val >> 8) & 255}, ${val & 255}, ${alpha})`;
};

export const Header: React.FunctionComponent<IHeaderProps> = (props) => {
    const { themeVariant } = props;
    // Safely handle opacity: strictly check for null/undefined, default to 100.
    const opacity = props.opacity ?? 100;

    // Default to a standard blue if theme is missing, but theme should be present
    const primaryColor = themeVariant?.palette?.themePrimary || '#0078d4';

    // Calculate alpha (0-1) from opacity (0-100)
    const alpha = opacity / 100;

    const backgroundStyle = {
        background: hexToRgba(primaryColor, alpha),
    };

    return (
        <div className={styles.header} style={backgroundStyle}>
            <h1 style={{ fontSize: props.titleFontSize || '24px', fontWeight: 600, margin: 0, color: '#ffffff' }}>
                {props.title || "📊 SharePoint Permission Viewer"}
            </h1>
        </div>
    );
};
