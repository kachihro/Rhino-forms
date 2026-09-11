import { IList } from '@pnp/sp/lists';
import '@pnp/sp/items';
import { IRhinoConfig, IFieldDefinition, IFilterState, IGridItem } from '../types/RhinoConfig';
import { systemFields } from '../config/helpers';

export interface IQueryOptions {
  filters: IFilterState;
  sortField: string;
  sortAsc: boolean;
  top: number;
  skip: number;
}

export interface IPage {
  items: IGridItem[];
  hasMore: boolean;
}

/**
 * CRUD against a resolved IList, driven entirely by the config's field list.
 * Uses plain arrays for multi-value fields (PnPjs v3 / JSON-light).
 */
export class GridDataService {
  constructor(private list: IList) {}

  /** Fields to select: every configured field plus the system columns (for the form footer / grid). */
  public static fieldsToLoad(config: IRhinoConfig): IFieldDefinition[] {
    const own = config.fields.filter(f => !f.isSystem);
    return own.concat(systemFields());
  }

  public async getPage(config: IRhinoConfig, opts: IQueryOptions): Promise<IPage> {
    const fields = GridDataService.fieldsToLoad(config);
    const select: string[] = ['Id'];
    const expand: string[] = [];
    fields.forEach(f => {
      if (f.internalName === 'ID') return;
      if (f.type === 'User') {
        select.push(`${f.internalName}/Id`, `${f.internalName}/Title`, `${f.internalName}/EMail`);
        expand.push(f.internalName);
      } else {
        select.push(f.internalName);
      }
    });

    let q = this.list.items.select(...select).top(opts.top + 1).skip(opts.skip);
    if (expand.length) q = q.expand(...expand);
    let sortField = opts.sortField === 'Id' ? 'ID' : (opts.sortField || 'ID');
    const sortDef = fields.filter(f => f.internalName === sortField)[0];
    if (sortDef && sortDef.type === 'User') sortField = `${sortField}/Title`;
    q = q.orderBy(sortField, opts.sortAsc);
    const filter = buildFilterString(config.fields.concat(systemFields()), opts.filters);
    if (filter) q = q.filter(filter);

    const rows = await q() as IGridItem[];
    const hasMore = rows.length > opts.top;
    return { items: hasMore ? rows.slice(0, opts.top) : rows, hasMore };
  }

  public async getById(config: IRhinoConfig, id: number): Promise<IGridItem> {
    const fields = GridDataService.fieldsToLoad(config);
    const select: string[] = ['Id'];
    const expand: string[] = [];
    fields.forEach(f => {
      if (f.internalName === 'ID') return;
      if (f.type === 'User') {
        select.push(`${f.internalName}/Id`, `${f.internalName}/Title`, `${f.internalName}/EMail`);
        expand.push(f.internalName);
      } else select.push(f.internalName);
    });
    let q = this.list.items.getById(id).select(...select);
    if (expand.length) q = q.expand(...expand);
    return await q() as IGridItem;
  }

  public async add(data: Record<string, unknown>): Promise<number> {
    const r = await this.list.items.add(data);
    return (r as any)?.data?.Id ?? (r as any)?.Id ?? 0;
  }

  public async update(id: number, data: Record<string, unknown>): Promise<void> {
    await this.list.items.getById(id).update(data);
  }

  public async delete(id: number): Promise<void> {
    await this.list.items.getById(id).recycle();
  }
}

function odataString(v: string): string {
  return `'${v.replace(/'/g, "''")}'`;
}

export function buildFilterString(fields: IFieldDefinition[], filters: IFilterState): string {
  const parts: string[] = [];
  fields.forEach(field => {
    const fv = filters[field.internalName];
    if (!fv) return;
    const n = field.internalName;
    switch (field.type) {
      case 'Text':
      case 'Note':
        if (fv.text) parts.push(`substringof(${odataString(fv.text)}, ${n})`);
        break;
      case 'Choice':
        if (fv.choice) parts.push(`${n} eq ${odataString(fv.choice)}`);
        break;
      case 'MultiChoice':
        if (fv.choice) parts.push(`substringof(${odataString(fv.choice)}, ${n})`);
        break;
      case 'DateTime':
        if (fv.dateFrom) parts.push(`${n} ge datetime'${fv.dateFrom}T00:00:00Z'`);
        if (fv.dateTo) parts.push(`${n} le datetime'${fv.dateTo}T23:59:59Z'`);
        break;
      case 'Number':
        if (fv.numberMin !== undefined) parts.push(`${n} ge ${fv.numberMin}`);
        if (fv.numberMax !== undefined) parts.push(`${n} le ${fv.numberMax}`);
        break;
      case 'Boolean':
        if (fv.boolValue !== undefined) parts.push(`${n} eq ${fv.boolValue ? 1 : 0}`);
        break;
      case 'User':
        if (fv.userTitle) parts.push(`substringof(${odataString(fv.userTitle)}, ${n}/Title)`);
        break;
      default:
        break;
    }
  });
  return parts.join(' and ');
}
