import {
  IRhinoConfig,
  IFieldDefinition,
  IFormConfig,
  IGridConfig,
  IGridColumn,
  IFormSection,
  FieldType,
  CellRenderer
} from '../types/RhinoConfig';

export const INTERNAL_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_]{0,31}$/;

/** Reserved SharePoint internal names that must not be re-provisioned. */
export const RESERVED_INTERNAL_NAMES = [
  'ID', 'Id', 'Created', 'Modified', 'Author', 'Editor', 'ContentType', 'Attachments',
  'GUID', 'FileRef', 'FileLeafRef', 'UniqueId', 'Order', 'WorkflowVersion'
];

/**
 * Turn "Due Date (est.)" into "DueDate" – a safe SharePoint internal name.
 * Never produces spaces or the dreaded `_x0020_` encoding.
 */
export function toInternalName(displayName: string): string {
  const cleaned = (displayName || '')
    .replace(/[^A-Za-z0-9 ]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join('');
  let name = cleaned.replace(/^[^A-Za-z]+/, '');
  if (!name) name = 'Field';
  return name.slice(0, 32);
}

/** Turn "Project Tracker" into "ProjectTracker" for the list URL segment. */
export function toListInternalName(title: string): string {
  return toInternalName(title).slice(0, 64);
}

export function isValidInternalName(name: string): boolean {
  return INTERNAL_NAME_PATTERN.test(name || '');
}

export function pluralise(label: string): string {
  if (!label) return 'items';
  if (/[^aeiou]y$/i.test(label)) return label.slice(0, -1) + 'ies';
  if (/(s|x|z|ch|sh)$/i.test(label)) return label + 'es';
  return label + 's';
}

export function defaultRenderer(type: FieldType): CellRenderer {
  switch (type) {
    case 'Choice': return 'badge';
    case 'MultiChoice': return 'tags';
    case 'User': return 'persona';
    case 'DateTime': return 'date';
    case 'Number': return 'number';
    case 'Boolean': return 'check';
    case 'URL': return 'link';
    default: return 'text';
  }
}

export function defaultColumnWidth(type: FieldType): number {
  switch (type) {
    case 'User': return 170;
    case 'Note': return 220;
    case 'DateTime': return 120;
    case 'Boolean': return 80;
    case 'Number': return 100;
    case 'MultiChoice': return 180;
    default: return 130;
  }
}

/** System columns that can be shown read-only in a grid / form. */
export function systemFields(): IFieldDefinition[] {
  return [
    { internalName: 'ID', displayName: 'ID', type: 'Number', numberFormat: 'integer', isSystem: true },
    { internalName: 'Created', displayName: 'Created', type: 'DateTime', dateOnly: false, isSystem: true },
    { internalName: 'Author', displayName: 'Created By', type: 'User', isSystem: true },
    { internalName: 'Modified', displayName: 'Modified', type: 'DateTime', dateOnly: false, isSystem: true },
    { internalName: 'Editor', displayName: 'Modified By', type: 'User', isSystem: true }
  ];
}

export function isFilterableType(type: FieldType): boolean {
  return type !== 'URL';
}

export function isSortableType(type: FieldType): boolean {
  return type !== 'URL' && type !== 'MultiChoice' && type !== 'Note';
}

export function buildDefaultColumn(field: IFieldDefinition): IGridColumn {
  return {
    field: field.internalName,
    width: defaultColumnWidth(field.type),
    sortable: field.sortable !== undefined ? field.sortable : isSortableType(field.type),
    filterable: field.filterable !== undefined ? field.filterable : (isFilterableType(field.type) && field.type !== 'Note'),
    render: 'auto'
  };
}

export function buildDefaultGrid(fields: IFieldDefinition[]): IGridConfig {
  return {
    columns: fields.filter(f => !f.isSystem).map(buildDefaultColumn),
    defaultSort: { field: 'ID', direction: 'desc' },
    pageSize: 50,
    allowSearch: true,
    allowFilter: true,
    allowSort: true,
    allowExport: true,
    allowAdd: true,
    allowEdit: true,
    allowDelete: true,
    allowColumnResize: true,
    density: 'comfortable',
    rowClick: 'edit',
    showCount: true
  };
}

export function newSectionId(): string {
  return 's' + Math.random().toString(36).slice(2, 8);
}

export function buildDefaultForm(fields: IFieldDefinition[]): IFormConfig {
  const section: IFormSection = {
    id: 'main',
    title: '',
    columns: 2,
    fields: fields
      .filter(f => !f.isSystem)
      .map(f => ({ field: f.internalName, width: f.type === 'Note' ? 'full' : 'half' }))
  };
  return {
    surface: 'panel',
    size: 'medium',
    sections: [section],
    showSystemInfo: true,
    addTitle: 'New {item}',
    editTitle: 'Edit {Title}',
    saveLabel: 'Save',
    cancelLabel: 'Cancel'
  };
}

export function createEmptyConfig(): IRhinoConfig {
  const fields: IFieldDefinition[] = [
    { internalName: 'Title', displayName: 'Title', type: 'Text', required: true }
  ];
  return {
    schemaVersion: 2,
    list: { title: '', internalName: '', description: '' },
    branding: { icon: '🦏', itemLabel: 'item', itemLabelPlural: 'items' },
    fields,
    form: buildDefaultForm(fields),
    grid: buildDefaultGrid(fields)
  };
}

export function sampleConfig(): IRhinoConfig {
  const fields: IFieldDefinition[] = [
    { internalName: 'Title', displayName: 'Project Name', type: 'Text', required: true, description: 'Short, unique project name' },
    { internalName: 'Status', displayName: 'Status', type: 'Choice', choices: ['Not Started', 'In Progress', 'On Hold', 'Complete'], defaultValue: 'Not Started', required: true },
    { internalName: 'Priority', displayName: 'Priority', type: 'Choice', choices: ['High', 'Medium', 'Low'], defaultValue: 'Medium' },
    { internalName: 'Owner', displayName: 'Owner', type: 'User' },
    { internalName: 'DueDate', displayName: 'Due Date', type: 'DateTime', dateOnly: true },
    { internalName: 'Budget', displayName: 'Budget', type: 'Number', numberFormat: 'currency', decimals: 2 },
    { internalName: 'Tags', displayName: 'Tags', type: 'MultiChoice', choices: ['Finance', 'Operations', 'IT', 'HR'] },
    { internalName: 'IsBillable', displayName: 'Billable', type: 'Boolean', defaultValue: true },
    { internalName: 'Notes', displayName: 'Notes', type: 'Note' }
  ];
  const cfg: IRhinoConfig = {
    schemaVersion: 2,
    list: { title: 'Project Tracker', internalName: 'ProjectTracker', description: 'Tracks projects and assignments' },
    branding: { title: 'Project Tracker', subtitle: 'All active and archived projects', icon: '🦏', itemLabel: 'project', itemLabelPlural: 'projects' },
    fields,
    form: {
      surface: 'panel',
      size: 'medium',
      showSystemInfo: true,
      addTitle: 'New {item}',
      editTitle: 'Edit {Title}',
      sections: [
        {
          id: 'details', title: 'Details', columns: 2,
          fields: [
            { field: 'Title', width: 'full', placeholder: 'e.g. Website refresh' },
            { field: 'Status', width: 'half' },
            { field: 'Priority', width: 'half' },
            { field: 'Owner', width: 'half' },
            { field: 'DueDate', width: 'half' }
          ]
        },
        {
          id: 'finance', title: 'Finance & tags', columns: 2,
          fields: [
            { field: 'Budget', width: 'half' },
            { field: 'IsBillable', width: 'half' },
            { field: 'Tags', width: 'full' }
          ]
        },
        {
          id: 'notes', title: 'Notes', columns: 1, collapsible: true,
          fields: [{ field: 'Notes', width: 'full' }]
        }
      ]
    },
    grid: buildDefaultGrid(fields)
  };
  cfg.grid.defaultSort = { field: 'DueDate', direction: 'asc' };
  cfg.grid.columns = cfg.grid.columns.filter(c => c.field !== 'Notes');
  return cfg;
}

export function cloneConfig<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function fieldByName(config: IRhinoConfig, name: string): IFieldDefinition | undefined {
  const all = config.fields.concat(systemFields());
  for (let i = 0; i < all.length; i++) {
    if (all[i].internalName === name) return all[i];
  }
  return undefined;
}

export function itemLabels(config: IRhinoConfig): { singular: string; plural: string } {
  const singular = config.branding.itemLabel || 'item';
  const plural = config.branding.itemLabelPlural || pluralise(singular);
  return { singular, plural };
}

export function configFileName(config: IRhinoConfig): string {
  const base = config.list.internalName || toListInternalName(config.list.title) || 'rhino-config';
  return `${base}.json`;
}
