import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneSlider,
    PropertyPaneToggle,
    PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import PermissionViewer from './components/PermissionViewer';
import {
    ThemeProvider,
    ThemeChangedEventArgs,
    IReadonlyTheme
} from '@microsoft/sp-component-base';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

export interface IPermissionViewerWebPartProps {
    description: string;
    headerOpacity: number;
    showStats: boolean;
    excludedLists: string[];
    themeVariant: IReadonlyTheme | undefined;

    buttonFontSize: string;
    showComponentHeader: boolean;
    webPartTitle: string;
    webPartTitleFontSize: string;
    contentFontSize: string;
}

export default class PermissionViewerWebPart extends BaseClientSideWebPart<IPermissionViewerWebPartProps> {

    private _themeProvider!: ThemeProvider;
    private _themeVariant: IReadonlyTheme | undefined;

    protected onInit(): Promise<void> {
        // Consume the new ThemeProvider service
        this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

        // If it exists, get the theme variant
        this._themeVariant = this._themeProvider.tryGetTheme();

        // Register a handler to be notified if the theme variant changes
        this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

        return super.onInit();
    }

    /**
     * Update the current theme variant reference and re-render.
     *
     * @param args The new theme
     */
    private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
        this._themeVariant = args.theme;
        this.render();
    }

    public render(): void {
        const element: React.ReactElement = React.createElement(
            PermissionViewer,
            {
                spHttpClient: this.context.spHttpClient,
                webUrl: this.context.pageContext.web.absoluteUrl,
                themeVariant: this._themeVariant,
                headerOpacity: this.properties.headerOpacity,
                showStats: this.properties.showStats,
                excludedLists: this.properties.excludedLists,

                buttonFontSize: this.properties.buttonFontSize,
                showComponentHeader: this.properties.showComponentHeader,
                webPartTitle: this.properties.webPartTitle,
                webPartTitleFontSize: this.properties.webPartTitleFontSize,
                contentFontSize: this.properties.contentFontSize || '14px'
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: "Permission Viewer Configuration"
                    },
                    groups: [
                        {
                            groupName: "Settings",
                            groupFields: [
                                PropertyPaneTextField('webPartTitle', {
                                    label: "Web Part Title"
                                }),
                                PropertyPaneDropdown('webPartTitleFontSize', {
                                    label: "Web Part Title Font Size",
                                    options: [
                                        { key: '20px', text: 'Medium (20px)' },
                                        { key: '24px', text: 'Large (24px)' },
                                        { key: '28px', text: 'Extra Large (28px)' },
                                        { key: '32px', text: 'Huge (32px)' }
                                    ]
                                }),
                                PropertyPaneDropdown('contentFontSize', {
                                    label: "Content Font Size",
                                    options: [
                                        { key: '12px', text: 'Small (12px)' },
                                        { key: '14px', text: 'Medium (14px)' },
                                        { key: '16px', text: 'Large (16px)' },
                                        { key: '18px', text: 'Extra Large (18px)' }
                                    ]
                                }),
                                PropertyPaneTextField('description', {
                                    label: "Description"
                                }),
                                PropertyPaneSlider('headerOpacity', {
                                    label: "Header Opacity",
                                    min: 0,
                                    max: 100,
                                    value: 100
                                }),
                                PropertyPaneToggle('showStats', {
                                    label: "Show Statistics Cards",
                                    checked: true
                                }),
                                PropertyFieldMultiSelect('excludedLists', {
                                    key: 'excludedLists',
                                    label: "Select System Lists to Exclude",
                                    options: [
                                        { key: "Site Assets", text: "Site Assets" },
                                        { key: "Site Pages", text: "Site Pages" },
                                        { key: "Style Library", text: "Style Library" },
                                        { key: "Master Page Gallery", text: "Master Page Gallery" },
                                        { key: "Form Templates", text: "Form Templates" },
                                        { key: "User Information List", text: "User Information List" },
                                        { key: "Composed Looks", text: "Composed Looks" },
                                        { key: "Solution Gallery", text: "Solution Gallery" },
                                        { key: "TaxonomyHiddenList", text: "TaxonomyHiddenList" },
                                        { key: "Appdata", text: "Appdata" },
                                        { key: "Appfiles", text: "Appfiles" }
                                    ],
                                    selectedKeys: this.properties.excludedLists
                                }),
                                PropertyPaneToggle('showComponentHeader', {
                                    label: "Show Web Part Header",
                                    checked: true
                                }),
                                PropertyPaneDropdown('buttonFontSize', {
                                    label: "Button Font Size",
                                    options: [
                                        { key: '10px', text: 'Small (10px)' },
                                        { key: '12px', text: 'Medium (12px)' },
                                        { key: '14px', text: 'Large (14px)' },
                                        { key: '16px', text: 'Extra Large (16px)' },
                                        { key: '18px', text: 'Huge (18px)' }
                                    ]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
