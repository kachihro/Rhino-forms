import { SPFI } from '@pnp/sp';
import { IList } from '@pnp/sp/lists';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import { IRhinoConfig, IFieldDefinition, IListRef, FieldType } from '../types/RhinoConfig';
import { toListInternalName, RESERVED_INTERNAL_NAMES } from '../config/helpers';

export interface IListSummary {
  id: string;
  title: string;
  internalName: string;
  serverRelativeUrl: string;
  description: string;
  itemCount: number;
}

export interface IProvisionReport {
  listCreated: boolean;
  listUrl: string;
  created: string[];
  updated: string[];
  skipped: string[];
  warnings: string[];
}

/** AddFieldOptions flags used by CreateFieldAsXml. */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;

export class ListSetupService {
  private webUrlPromise: Promise<string> | undefined;

  constructor(private sp: SPFI) {}

  /** Server-relative URL of the current web, cached. */
  public async getWebServerRelativeUrl(): Promise<string> {
    if (!this.webUrlPromise) {
      this.webUrlPromise = this.sp.web.select('ServerRelativeUrl')().then(w => w.ServerRelativeUrl.replace(/\/$/, ''));
    }
    return this.webUrlPromise;
  }

  /** All non-hidden custom lists (generic list template) in the current web. */
  public async getLists(): Promise<IListSummary[]> {
    const lists = await this.sp.web.lists
      .select('Id', 'Title', 'Description', 'ItemCount', 'Hidden', 'BaseTemplate', 'RootFolder/ServerRelativeUrl', 'RootFolder/Name')
      .expand('RootFolder')
      .filter('Hidden eq false and BaseTemplate eq 100')
      .orderBy('Title')();
    return lists.map((l: any) => ({
      id: l.Id,
      title: l.Title,
      internalName: l.RootFolder?.Name || '',
      serverRelativeUrl: l.RootFolder?.ServerRelativeUrl || '',
      description: l.Description || '',
      itemCount: l.ItemCount || 0
    }));
  }

  /**
   * Resolve a list by URL segment first (rename-safe), then by title.
   * Throws if neither works.
   */
  public async resolveList(ref: IListRef): Promise<IList> {
    if (ref.internalName) {
      const web = await this.getWebServerRelativeUrl();
      const byUrl = this.sp.web.getList(`${web}/Lists/${ref.internalName}`);
      try {
        await byUrl.select('Id')();
        return byUrl;
      } catch {
        /* fall through to title */
      }
    }
    const byTitle = this.sp.web.lists.getByTitle(ref.title);
    await byTitle.select('Id')();
    return byTitle;
  }

  public async listExists(ref: IListRef): Promise<boolean> {
    try {
      await this.resolveList(ref);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Create the list (if needed) and bring its columns in line with the config.
   * Idempotent: safe to re-run after editing the config.
   */
  public async provisionList(config: IRhinoConfig): Promise<IProvisionReport> {
    const report: IProvisionReport = { listCreated: false, listUrl: '', created: [], updated: [], skipped: [], warnings: [] };
    const ref = config.list;
    const internalName = ref.internalName || toListInternalName(ref.title);
    const webUrl = await this.getWebServerRelativeUrl();

    let list: IList;
    const exists = await this.listExists(ref);
    if (!exists) {
      // Create with the *internal* name so the URL is /Lists/<internalName>,
      // then rename the display title. This is the only way to control the URL.
      await this.sp.web.lists.add(internalName, ref.description || '', 100, false);
      list = this.sp.web.lists.getByTitle(internalName);
      if (ref.title && ref.title !== internalName) {
        await list.update({ Title: ref.title });
      }
      report.listCreated = true;
    } else {
      list = await this.resolveList(ref);
      if (ref.description !== undefined) {
        const current = await list.select('Description', 'Title')();
        if ((current.Description || '') !== (ref.description || '') || current.Title !== ref.title) {
          await list.update({ Title: ref.title, Description: ref.description || '' });
        }
      }
    }
    const info = await list.select('RootFolder/ServerRelativeUrl').expand('RootFolder')();
    report.listUrl = (info as any).RootFolder?.ServerRelativeUrl || `${webUrl}/Lists/${internalName}`;

    const existing = await list.fields
      .select('InternalName', 'Title', 'TypeAsString', 'Required', 'Description', 'Choices', 'Hidden')();
    const byInternal: Record<string, any> = {};
    existing.forEach((f: any) => { byInternal[f.InternalName.toLowerCase()] = f; });

    for (const field of config.fields) {
      if (field.isSystem) continue;
      const key = field.internalName.toLowerCase();
      const current = byInternal[key];

      if (key === 'title') {
        const changes: Record<string, unknown> = {};
        if (current && field.displayName && current.Title !== field.displayName) changes.Title = field.displayName;
        if (current && (current.Description || '') !== (field.description || '')) changes.Description = field.description || '';
        if (Object.keys(changes).length > 0) {
          await list.fields.getByInternalNameOrTitle('Title').update(changes);
          report.updated.push('Title');
        } else {
          report.skipped.push('Title');
        }
        continue;
      }

      if (RESERVED_INTERNAL_NAMES.indexOf(field.internalName) >= 0) {
        report.warnings.push(`"${field.internalName}" is a built-in column and was not created.`);
        continue;
      }

      if (!current) {
        await this.createField(list, field);
        report.created.push(field.internalName);
        continue;
      }

      // Existing column: sync the safe-to-change properties.
      const spType = this.toSpType(field.type);
      if (current.TypeAsString !== spType && !(current.TypeAsString === 'UserMulti' && spType === 'User')) {
        report.warnings.push(`"${field.internalName}" already exists as ${current.TypeAsString}; config says ${spType}. Type was left unchanged.`);
      }
      const changes: Record<string, unknown> = {};
      if (current.Title !== field.displayName) changes.Title = field.displayName;
      if ((current.Description || '') !== (field.description || '')) changes.Description = field.description || '';
      if (!!current.Required !== !!field.required) changes.Required = !!field.required;
      let fieldTypeName: string | undefined;
      if ((field.type === 'Choice' || field.type === 'MultiChoice') && field.choices) {
        const oldChoices: string[] = current.Choices || [];
        if (JSON.stringify(oldChoices) !== JSON.stringify(field.choices)) {
          changes.Choices = field.choices;
          fieldTypeName = field.type === 'Choice' ? 'SP.FieldChoice' : 'SP.FieldMultiChoice';
        }
      }
      if (Object.keys(changes).length > 0) {
        await list.fields.getByInternalNameOrTitle(field.internalName).update(changes, fieldTypeName);
        report.updated.push(field.internalName);
      } else {
        report.skipped.push(field.internalName);
      }
    }

    return report;
  }

  /** Read an existing list's columns into field definitions (for "start from existing list"). */
  public async importFields(ref: IListRef): Promise<IFieldDefinition[]> {
    const list = await this.resolveList(ref);
    const spFields = await list.fields
      .select('InternalName', 'Title', 'TypeAsString', 'Required', 'Description', 'Choices', 'Hidden',
        'ReadOnlyField', 'MaxLength', 'DefaultValue', 'AllowMultipleValues', 'DisplayFormat', 'CanBeDeleted', 'FromBaseType', 'Sealed')
      .filter('Hidden eq false')();

    const out: IFieldDefinition[] = [];
    for (const f of spFields as any[]) {
      if (f.InternalName === 'Title') {
        out.unshift({ internalName: 'Title', displayName: f.Title, type: 'Text', required: true, description: f.Description || undefined });
        continue;
      }
      if (f.ReadOnlyField || f.Sealed) continue;
      if (f.FromBaseType && f.InternalName !== 'Title') continue; // ContentType, Attachments, etc.
      if (f.InternalName === 'ContentType' || f.InternalName === 'Attachments') continue;
      const mapped = this.fromSpType(f.TypeAsString);
      if (!mapped) continue;

      const def: IFieldDefinition = {
        internalName: f.InternalName,
        displayName: f.Title,
        type: mapped,
        required: !!f.Required,
        description: f.Description || undefined
      };
      if (mapped === 'Choice' || mapped === 'MultiChoice') def.choices = f.Choices || [];
      if (mapped === 'Text' && f.MaxLength) def.maxLength = f.MaxLength;
      if (mapped === 'DateTime') def.dateOnly = f.DisplayFormat === 0;
      if (mapped === 'User') def.allowMultiple = f.TypeAsString === 'UserMulti' || !!f.AllowMultipleValues;
      if (mapped === 'Number' && f.TypeAsString === 'Currency') def.numberFormat = 'currency';
      if (f.DefaultValue !== null && f.DefaultValue !== undefined && f.DefaultValue !== '') {
        def.defaultValue = mapped === 'Boolean' ? f.DefaultValue === '1' : mapped === 'Number' ? Number(f.DefaultValue) : f.DefaultValue;
      }
      out.push(def);
    }
    return out;
  }

  private async createField(list: IList, field: IFieldDefinition): Promise<void> {
    const xml = this.buildSchemaXml(field);
    await list.fields.createFieldAsXml({
      SchemaXml: xml,
      Options: ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW
    });
  }

  private toSpType(type: FieldType): string {
    switch (type) {
      case 'User': return 'User';
      default: return type;
    }
  }

  private fromSpType(spType: string): FieldType | undefined {
    switch (spType) {
      case 'Text': return 'Text';
      case 'Note': return 'Note';
      case 'Choice': return 'Choice';
      case 'MultiChoice': return 'MultiChoice';
      case 'DateTime': return 'DateTime';
      case 'Number':
      case 'Currency':
      case 'Integer': return 'Number';
      case 'Boolean': return 'Boolean';
      case 'User':
      case 'UserMulti': return 'User';
      case 'URL': return 'URL';
      default: return undefined;
    }
  }

  private buildSchemaXml(field: IFieldDefinition): string {
    const name = this.escapeXml(field.internalName);
    const display = this.escapeXml(field.displayName || field.internalName);
    const required = field.required ? 'TRUE' : 'FALSE';
    const description = field.description ? ` Description="${this.escapeXml(field.description)}"` : '';
    const common = `Name="${name}" StaticName="${name}" DisplayName="${display}" Required="${required}"${description}`;
    const defaultXml = (v: unknown): string => (v === undefined || v === null || v === '') ? '' : `<Default>${this.escapeXml(String(v))}</Default>`;

    switch (field.type) {
      case 'Text':
        return `<Field Type="Text" ${common} MaxLength="${field.maxLength || 255}">${defaultXml(field.defaultValue)}</Field>`;
      case 'Note':
        return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" AppendOnly="FALSE"/>`;
      case 'Choice':
      case 'MultiChoice': {
        const choices = (field.choices || []).map(c => `<CHOICE>${this.escapeXml(c)}</CHOICE>`).join('');
        const format = field.type === 'Choice' ? ' Format="Dropdown"' : '';
        return `<Field Type="${field.type}" ${common}${format}>${defaultXml(field.defaultValue)}<CHOICES>${choices}</CHOICES></Field>`;
      }
      case 'DateTime': {
        const fmt = field.dateOnly === false ? 'DateTime' : 'DateOnly';
        return `<Field Type="DateTime" ${common} Format="${fmt}" FriendlyDisplayFormat="Disabled"/>`;
      }
      case 'Number': {
        const decimals = field.numberFormat === 'integer' ? 0 : (field.decimals !== undefined ? field.decimals : (field.numberFormat === 'currency' ? 2 : undefined));
        const dec = decimals !== undefined ? ` Decimals="${decimals}"` : '';
        const pct = field.numberFormat === 'percent' ? ' Percentage="TRUE"' : '';
        const min = field.min !== undefined ? ` Min="${field.min}"` : '';
        const max = field.max !== undefined ? ` Max="${field.max}"` : '';
        if (field.numberFormat === 'currency') {
          return `<Field Type="Currency" ${common}${dec}${min}${max} LCID="3081">${defaultXml(field.defaultValue)}</Field>`;
        }
        return `<Field Type="Number" ${common}${dec}${pct}${min}${max}>${defaultXml(field.defaultValue)}</Field>`;
      }
      case 'Boolean':
        return `<Field Type="Boolean" ${common}><Default>${field.defaultValue ? 1 : 0}</Default></Field>`;
      case 'User': {
        const type = field.allowMultiple ? 'UserMulti' : 'User';
        const mult = field.allowMultiple ? ' Mult="TRUE"' : '';
        return `<Field Type="${type}" ${common} UserSelectionMode="PeopleOnly" ShowField="ImnName" List="UserInfo"${mult}/>`;
      }
      case 'URL':
        return `<Field Type="URL" ${common} Format="Hyperlink"/>`;
      default:
        return `<Field Type="Text" ${common} MaxLength="255"/>`;
    }
  }

  private escapeXml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
