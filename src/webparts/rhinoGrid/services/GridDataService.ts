import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IFieldDefinition, IFilterState, IGridItem } from '../types/IFieldConfig';

export class GridDataService {
  constructor(private sp: SPFI) {}

  public async getItems(
    listName: string,
    fields: IFieldDefinition[],
    filters: IFilterState,
    sortField: string = 'ID',
    sortAsc: boolean = true,
    top: number = 100,
    skip: number = 0
  ): Promise<IGridItem[]> {
    const selectFields: string[] = ['Id'];
    const expandFields: string[] = [];

    for (const field of fields) {
      if (field.type === 'User') {
        selectFields.push(`${field.internalName}/Id`, `${field.internalName}/Title`);
        expandFields.push(field.internalName);
      } else {
        selectFields.push(field.internalName);
      }
    }

    const filterStr = this.buildFilterString(fields, filters);

    let query = this.sp.web.lists
      .getByTitle(listName)
      .items
      .select(...selectFields)
      .orderBy(sortField, sortAsc)
      .top(top)
      .skip(skip);

    if (expandFields.length > 0) {
      query = query.expand(...expandFields);
    }

    if (filterStr) {
      query = query.filter(filterStr);
    }

    return await query() as IGridItem[];
  }

  public async addItem(listName: string, data: Record<string, unknown>): Promise<IGridItem> {
    const result = await this.sp.web.lists.getByTitle(listName).items.add(data);
    return result as unknown as IGridItem;
  }

  public async updateItem(listName: string, id: number, data: Record<string, unknown>): Promise<void> {
    await this.sp.web.lists.getByTitle(listName).items.getById(id).update(data);
  }

  public async deleteItem(listName: string, id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(listName).items.getById(id).delete();
  }

  private buildFilterString(fields: IFieldDefinition[], filters: IFilterState): string {
    const parts: string[] = [];

    for (const field of fields) {
      const filterValue = filters[field.internalName];
      if (!filterValue) continue;

      switch (field.type) {
        case 'Text':
        case 'Note':
          if (filterValue.text) {
            const escaped = filterValue.text.replace(/'/g, "''");
            parts.push(`substringof('${escaped}', ${field.internalName})`);
          }
          break;
        case 'Choice':
          if (filterValue.choice) {
            const escaped = filterValue.choice.replace(/'/g, "''");
            parts.push(`${field.internalName} eq '${escaped}'`);
          }
          break;
        case 'DateTime':
          if (filterValue.dateFrom) {
            parts.push(`${field.internalName} ge datetime'${filterValue.dateFrom}T00:00:00Z'`);
          }
          if (filterValue.dateTo) {
            parts.push(`${field.internalName} le datetime'${filterValue.dateTo}T23:59:59Z'`);
          }
          break;
        case 'Number':
          if (filterValue.numberMin !== undefined) {
            parts.push(`${field.internalName} ge ${filterValue.numberMin}`);
          }
          if (filterValue.numberMax !== undefined) {
            parts.push(`${field.internalName} le ${filterValue.numberMax}`);
          }
          break;
        case 'Boolean':
          if (filterValue.boolValue !== undefined) {
            parts.push(`${field.internalName} eq ${filterValue.boolValue ? '1' : '0'}`);
          }
          break;
        default:
          break;
      }
    }

    return parts.join(' and ');
  }
}
