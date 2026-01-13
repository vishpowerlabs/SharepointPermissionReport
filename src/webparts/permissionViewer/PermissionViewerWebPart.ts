import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneSlider,
    PropertyPaneToggle
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
                excludedLists: this.properties.excludedLists
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
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
