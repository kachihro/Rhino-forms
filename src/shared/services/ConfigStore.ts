import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import { IRhinoConfig } from '../types/RhinoConfig';
import { parseConfig, serializeConfig } from '../config/normalize';

export const CONFIG_FOLDER = 'RhinoForms';

export interface IStoredConfig {
  name: string;
  serverRelativeUrl: string;
  modified: string;
}

/**
 * Persists config JSON files under  <web>/SiteAssets/RhinoForms/<name>.json
 * so the designer and any number of list web parts share one source of truth.
 * Edit the file in place (or re-save from the designer) and every web part
 * pointing at it picks the change up on next load.
 */
export class ConfigStore {
  private folderUrlPromise: Promise<string> | undefined;

  constructor(private sp: SPFI) {}

  public async getFolderUrl(): Promise<string> {
    if (!this.folderUrlPromise) {
      this.folderUrlPromise = (async () => {
        const lib = await this.sp.web.lists.ensureSiteAssetsLibrary();
        const root = await lib.rootFolder.select('ServerRelativeUrl')();
        return `${root.ServerRelativeUrl.replace(/\/$/, '')}/${CONFIG_FOLDER}`;
      })();
    }
    return this.folderUrlPromise;
  }

  public async ensureFolder(): Promise<string> {
    const url = await this.getFolderUrl();
    const folder = this.sp.web.getFolderByServerRelativePath(url);
    try {
      const info = await folder.select('Exists')();
      if (!info.Exists) throw new Error('missing');
    } catch {
      const parent = url.slice(0, url.lastIndexOf('/'));
      await this.sp.web.getFolderByServerRelativePath(parent).folders.addUsingPath(CONFIG_FOLDER);
    }
    return url;
  }

  public async list(): Promise<IStoredConfig[]> {
    try {
      const url = await this.getFolderUrl();
      const files = await this.sp.web.getFolderByServerRelativePath(url).files
        .select('Name', 'ServerRelativeUrl', 'TimeLastModified')
        .orderBy('Name')();
      return files
        .filter((f: any) => /\.json$/i.test(f.Name))
        .map((f: any) => ({ name: f.Name, serverRelativeUrl: f.ServerRelativeUrl, modified: f.TimeLastModified }));
    } catch {
      return [];
    }
  }

  public async load(serverRelativeUrl: string): Promise<IRhinoConfig> {
    const text = await this.sp.web.getFileByServerRelativePath(serverRelativeUrl).getText();
    return parseConfig(text);
  }

  public async loadText(serverRelativeUrl: string): Promise<string> {
    return this.sp.web.getFileByServerRelativePath(serverRelativeUrl).getText();
  }

  public async save(fileName: string, config: IRhinoConfig): Promise<IStoredConfig> {
    const url = await this.ensureFolder();
    const name = /\.json$/i.test(fileName) ? fileName : `${fileName}.json`;
    const content = serializeConfig(config);
    const result = await this.sp.web.getFolderByServerRelativePath(url).files.addUsingPath(name, content, { Overwrite: true });
    const data: any = (result as any).data || result;
    return { name, serverRelativeUrl: data?.ServerRelativeUrl || `${url}/${name}`, modified: data?.TimeLastModified || new Date().toISOString() };
  }
}
