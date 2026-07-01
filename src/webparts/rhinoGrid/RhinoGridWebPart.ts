import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';

import * as strings from 'RhinoGridWebPartStrings';
import RhinoGridApp, { IRhinoGridAppProps } from './components/RhinoGridApp';
import { IGridConfig } from './types/IFieldConfig';

export interface IRhinoGridWebPartProps {
  gridConfigJson: string;
  listName: string;
}

export default class RhinoGridWebPart extends BaseClientSideWebPart<IRhinoGridWebPartProps> {

  private _sp: SPFI;

  public async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  protected get isFullPage(): boolean {
    return true;
  }

  public render(): void {
    let config: IGridConfig | undefined;
    try {
      if (this.properties.gridConfigJson && this.properties.gridConfigJson.trim()) {
        config = JSON.parse(this.properties.gridConfigJson) as IGridConfig;
      }
    } catch {
      config = undefined;
    }

    const element: React.ReactElement<IRhinoGridAppProps> = React.createElement(
      RhinoGridApp,
      {
        sp: this._sp,
        config,
        displayMode: this.displayMode,
        onConfigSaved: (json: string, listName: string) => {
          this.properties.gridConfigJson = json;
          this.properties.listName = listName;
          this.render();
        }
      }
    );

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
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const hasConfig = !!(this.properties.gridConfigJson && this.properties.gridConfigJson.trim());

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneLabel('listName', {
                  text: hasConfig
                    ? `Connected list: ${this.properties.listName || 'Unknown'}`
                    : 'No configuration loaded. Use the setup screen to upload a JSON config file.'
                }),
                PropertyPaneTextField('gridConfigJson', {
                  label: strings.GridConfigJsonLabel,
                  description: 'Paste or edit your JSON field configuration here. Changes take effect when the page is saved.',
                  multiline: true,
                  rows: 12,
                  resizable: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
