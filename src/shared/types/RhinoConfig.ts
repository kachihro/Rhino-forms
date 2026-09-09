/**
 * Rhino Forms configuration schema (v2).
 *
 * One JSON document drives everything:
 *   - `list`      which SharePoint list to bind to (and how to provision it)
 *   - `fields`    the data schema (what columns exist and their types)
 *   - `form`      how the add/edit form is laid out
 *   - `grid`      how the list view behaves (columns, sort, filter, actions)
 *   - `branding`  header, icon, labels
 *
 * v1 documents (the original `{ listName, fields[] }` shape) are still accepted
 * and upgraded on load by `normalizeConfig()` in ../config/normalize.ts.
 */

export type FieldType =
  | 'Text'
  | 'Note'
  | 'Choice'
  | 'MultiChoice'
  | 'DateTime'
  | 'Number'
  | 'Boolean'
  | 'User'
  | 'URL';

export const FIELD_TYPES: FieldType[] = ['Text', 'Note', 'Choice', 'MultiChoice', 'DateTime', 'Number', 'Boolean', 'User', 'URL'];

export type NumberFormat = 'decimal' | 'integer' | 'currency' | 'percent';

export interface IFieldDefinition {
  /** SharePoint internal (static) name. Letters, digits and underscore only; must start with a letter. */
  internalName: string;
  /** Human-friendly column title. */
  displayName: string;
  type: FieldType;
  /** Shown as help text in the form and stored as the SharePoint field description. */
  description?: string;
  required?: boolean;
  /** Choice / MultiChoice options. */
  choices?: string[];
  /** Initial value for new items. */
  defaultValue?: unknown;
  /** Text only. Defaults to 255. */
  maxLength?: number;
  /** Number only. */
  numberFormat?: NumberFormat;
  decimals?: number;
  min?: number;
  max?: number;
  /** DateTime only. Defaults to true (date without time). */
  dateOnly?: boolean;
  /** User only: allow selecting several people. */
  allowMultiple?: boolean;
  /**
   * Built-in SharePoint column (ID, Created, Modified, Author, Editor).
   * Never provisioned, never editable; can still be shown in the grid / form.
   */
  isSystem?: boolean;

  /** @deprecated v1 – use grid.columns[].filterable. Still honoured on load. */
  filterable?: boolean;
  /** @deprecated v1 – use grid.columns[].sortable. Still honoured on load. */
  sortable?: boolean;
}

export type FieldWidth = 'full' | 'half' | 'third' | 'twoThirds';

export interface IFormFieldPlacement {
  /** internalName of the field. */
  field: string;
  width?: FieldWidth;
  /** Override the label shown on the form. */
  label?: string;
  placeholder?: string;
  /** Override the help text under the control. */
  helpText?: string;
  readOnly?: boolean;
}

export interface IFormSection {
  id: string;
  title?: string;
  description?: string;
  /** Column count for the section grid. Field widths are relative to the section. */
  columns?: 1 | 2 | 3;
  collapsible?: boolean;
  collapsedByDefault?: boolean;
  fields: IFormFieldPlacement[];
}

export type FormSurface = 'panel' | 'dialog';
export type FormSize = 'medium' | 'large' | 'extraLarge';

export interface IFormConfig {
  surface: FormSurface;
  size: FormSize;
  sections: IFormSection[];
  /** Show Created / Modified metadata when editing. */
  showSystemInfo?: boolean;
  /** Text shown in the header. `{item}` = itemLabel, `{Title}` = current title. */
  addTitle?: string;
  editTitle?: string;
  saveLabel?: string;
  cancelLabel?: string;
}

export type CellRenderer =
  | 'auto'
  | 'text'
  | 'badge'
  | 'persona'
  | 'date'
  | 'datetime'
  | 'number'
  | 'currency'
  | 'percent'
  | 'link'
  | 'check'
  | 'tags';

export type BadgeTone = 'neutral' | 'info' | 'success' | 'warning' | 'danger' | 'purple' | 'teal';

export interface IGridColumn {
  field: string;
  label?: string;
  /** Minimum width in px. */
  width?: number;
  maxWidth?: number;
  sortable?: boolean;
  filterable?: boolean;
  render?: CellRenderer;
  /** Choice value -> tone, used by the badge renderer. Unlisted values are auto-coloured. */
  badgeTones?: Record<string, BadgeTone>;
  align?: 'left' | 'center' | 'right';
}

export type GridDensity = 'compact' | 'comfortable';
export type RowClickAction = 'edit' | 'view' | 'none';

export interface IGridConfig {
  columns: IGridColumn[];
  defaultSort?: { field: string; direction: 'asc' | 'desc' };
  pageSize?: number;
  allowSearch?: boolean;
  allowFilter?: boolean;
  allowSort?: boolean;
  allowExport?: boolean;
  allowAdd?: boolean;
  allowEdit?: boolean;
  allowDelete?: boolean;
  allowColumnResize?: boolean;
  density?: GridDensity;
  rowClick?: RowClickAction;
  showCount?: boolean;
  emptyMessage?: string;
}

export interface IListRef {
  /** Display title of the list. */
  title: string;
  /**
   * URL segment of the list (…/Lists/<internalName>). Set at provisioning time.
   * When present the list is resolved by URL so renaming the list later is safe.
   */
  internalName?: string;
  description?: string;
}

export interface IBranding {
  /** Header title. Defaults to list.title. */
  title?: string;
  subtitle?: string;
  /** Emoji or Fluent icon name. */
  icon?: string;
  /** Accent colour (hex). */
  accent?: string;
  itemLabel?: string;
  itemLabelPlural?: string;
}

export interface IRhinoConfig {
  schemaVersion: 2;
  list: IListRef;
  branding: IBranding;
  fields: IFieldDefinition[];
  form: IFormConfig;
  grid: IGridConfig;
}

/* ------------------------------------------------------------------ */
/* Runtime types                                                       */
/* ------------------------------------------------------------------ */

export interface IUserValue {
  Id: number;
  Title?: string;
  EMail?: string;
  Name?: string;
}

export interface IUrlValue {
  Url?: string;
  Description?: string;
}

export interface IGridItem {
  Id: number;
  Title?: string;
  Created?: string;
  Modified?: string;
  Author?: IUserValue;
  Editor?: IUserValue;
  [key: string]: unknown;
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

/** v1 (legacy) shape kept for the normaliser. */
export interface ILegacyGridConfig {
  listName: string;
  listDescription?: string;
  icon?: string;
  itemLabel?: string;
  fields: IFieldDefinition[];
}
