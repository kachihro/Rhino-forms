export type FieldType = 'Text' | 'Note' | 'Choice' | 'DateTime' | 'Number' | 'Boolean' | 'User' | 'URL';

export interface IFieldDefinition {
  internalName: string;
  displayName: string;
  type: FieldType;
  required?: boolean;
  filterable?: boolean;
  sortable?: boolean;
  choices?: string[];
  maxLength?: number;
  defaultValue?: unknown;
}

export interface IGridConfig {
  listName: string;
  listDescription?: string;
  fields: IFieldDefinition[];
}

export interface IFilterValue {
  text?: string;
  choice?: string;
  dateFrom?: string;
  dateTo?: string;
  numberMin?: number;
  numberMax?: number;
  boolValue?: boolean;
  userTitle?: string;
}

export interface IFilterState {
  [fieldInternalName: string]: IFilterValue;
}

export interface IGridItem {
  Id: number;
  Title?: string;
  [key: string]: unknown;
}
