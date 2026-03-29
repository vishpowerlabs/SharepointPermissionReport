
/**
 * Inject custom global styles to override SharePoint/Workbench default max-widths.
 * This ensures the web part takes up the full width of the page.
 */
export const injectGlobalStyles = (): void => {
    try {
        if (typeof document === 'undefined' || !document.head) return;

        const styleId = 'permission-viewer-global-overrides';
        if (!document.getElementById(styleId)) {
            const styleElement = document.createElement('style');
            styleElement.id = styleId;
            styleElement.innerHTML = `
            .p_ZwSiC_hHQBj {
                max-width: 100% !important;
                width: 100% !important;
            }
            .CanvasComponent.LCS .CanvasZone {
                max-width: 100% !important;
                width: 100% !important;
                padding: 0 !important;
                margin: 0 auto !important;
            }
            .CanvasComponent {
                max-width: 100% !important;
                width: 100% !important;
            }
            .LCS {
                max-width: 100% !important;
                width: 100% !important;
            }
            .CanvasZone {
                max-width: 100% !important;
                width: 100% !important;
            }
        `;
            document.head.appendChild(styleElement);
        }
    } catch (e) {
        console.error("Error injecting global styles", e);
    }
};
