import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/profiles';
import '@pnp/sp/site-users/web';

import * as strings from 'RhinoDesignerWebPartStrings';
import DesignerApp, { IDesignerAppProps } from './components/DesignerApp';

export interface IRhinoDesignerWebPartProps {
  lastConfigFileUrl: string;
}

export default class RhinoDesignerWebPart extends BaseClientSideWebPart<IRhinoDesignerWebPartProps> {
  private _sp: SPFI;

  public async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  protected get isFullPage(): boolean { return true; }
  protected get isFullBleed(): boolean { return true; }

  public render(): void {
    const element: React.ReactElement<IDesignerAppProps> = React.createElement(DesignerApp, {
      sp: this._sp,
      webUrl: this.context.pageContext.web.absoluteUrl,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      draftKey: `rhino-designer-draft:${this.context.pageContext.web.id.toString()}`
    });
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
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneLabel('about', { text: 'Rhino List Designer – define, provision and configure lists for the Rhino List web part.' })
          ]
        }]
      }]
    };
  }
}
