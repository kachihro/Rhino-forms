import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

import * as strings from 'RhinoGridWebPartStrings';
import RhinoGridApp, { IRhinoGridAppProps } from './components/RhinoGridApp';
import { ConfigStore, IStoredConfig } from '../../shared/services/ConfigStore';

export type ConfigSource = 'file' | 'inline';

export interface IRhinoGridWebPartProps {
  /** Where the config comes from: a JSON file in Site Assets/RhinoForms, or JSON pasted into the property pane. */
  configSource: ConfigSource;
  /** Server-relative URL of the config file when configSource === 'file'. */
  configFileUrl: string;
  /** Inline JSON when configSource === 'inline'. */
  gridConfigJson: string;
  hideHeader: boolean;
  /** Legacy (v1) – kept so existing pages keep working. */
  listName: string;
}

export default class RhinoGridWebPart extends BaseClientSideWebPart<IRhinoGridWebPartProps> {
  private _sp: SPFI;
  private _store: ConfigStore;
  private _configFiles: IStoredConfig[] = [];
  private _filesLoaded = false;

  public async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
    this._store = new ConfigStore(this._sp);
    // Backwards compatibility: pages saved with v1 only had gridConfigJson.
    if (!this.properties.configSource) {
      this.properties.configSource = this.properties.gridConfigJson && this.properties.gridConfigJson.trim() ? 'inline' : 'file';
    }
  }

  protected get isFullPage(): boolean { return true; }
  protected get isFullBleed(): boolean { return true; }

  public render(): void {
    const element: React.ReactElement<IRhinoGridAppProps> = React.createElement(RhinoGridApp, {
      sp: this._sp,
      store: this._store,
      configSource: this.properties.configSource || 'file',
      configFileUrl: this.properties.configFileUrl || '',
      inlineJson: this.properties.gridConfigJson || '',
      hideHeader: !!this.properties.hideHeader,
      displayMode: this.displayMode,
      designerUrl: `${this.context.pageContext.web.absoluteUrl}/SitePages/RhinoDesigner.aspx`,
      onOpenPropertyPane: () => this.context.propertyPane.open(),
      onPickConfigFile: (url: string) => {
        this.properties.configSource = 'file';
        this.properties.configFileUrl = url;
        this.render();
      }
    });
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (this._filesLoaded) return;
    try {
      this._configFiles = await this._store.list();
    } catch {
      this._configFiles = [];
    }
    this._filesLoaded = true;
    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fileOptions: IPropertyPaneDropdownOption[] = this._configFiles.map(f => ({ key: f.serverRelativeUrl, text: f.name.replace(/\.json$/i, '') }));
    const source = this.properties.configSource || 'file';

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('configSource', {
                  label: 'Configuration source',
                  options: [
                    { key: 'file', text: 'Config file from Site Assets (recommended)', iconProps: { officeFabricIconFontName: 'Document' } },
                    { key: 'inline', text: 'JSON pasted below', iconProps: { officeFabricIconFontName: 'Code' } }
                  ]
                }),
                ...(source === 'file'
                  ? [
                    PropertyPaneDropdown('configFileUrl', {
                      label: 'Config file',
                      options: fileOptions,
                      disabled: !this._filesLoaded,
                      selectedKey: this.properties.configFileUrl
                    }),
                    PropertyPaneLabel('fileHint', {
                      text: this._filesLoaded && fileOptions.length === 0
                        ? 'No config files found in Site Assets/RhinoForms. Use the Rhino List Designer web part to create one.'
                        : 'Files are read from Site Assets/RhinoForms. Save from the Designer or edit the JSON file directly.'
                    })
                  ]
                  : [
                    PropertyPaneTextField('gridConfigJson', {
                      label: strings.GridConfigJsonLabel,
                      description: 'Paste a Rhino Forms JSON config. v1 configs are upgraded automatically.',
                      multiline: true,
                      rows: 14,
                      resizable: true
                    })
                  ]),
                PropertyPaneToggle('hideHeader', { label: 'Hide header card', onText: 'Hidden', offText: 'Shown' }),
                PropertyPaneLink('designerLink', {
                  text: 'Open the Rhino List Designer (new tab)',
                  href: `${this.context.pageContext.web.absoluteUrl}/SitePages/RhinoDesigner.aspx`,
                  target: '_blank'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
