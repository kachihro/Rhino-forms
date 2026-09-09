import {
  IRhinoConfig,
  ILegacyGridConfig,
  IFieldDefinition,
  IFormConfig,
  IGridConfig,
  IFormSection
} from '../types/RhinoConfig';
import {
  buildDefaultForm,
  buildDefaultGrid,
  buildDefaultColumn,
  toListInternalName,
  pluralise,
  systemFields,
  cloneConfig
} from './helpers';

/**
 * Accepts any JSON that looks like a Rhino config (v1 or v2, partial or complete)
 * and returns a fully-populated v2 config. Never throws on missing optional
 * sections – it fills them in from the field list so a hand-written minimal
 * config still "just works".
 */
export function normalizeConfig(raw: unknown): IRhinoConfig {
  if (!raw || typeof raw !== 'object') {
    throw new Error('Configuration must be a JSON object.');
  }
  const src = raw as Partial<IRhinoConfig> & Partial<ILegacyGridConfig>;

  // ---- fields ---------------------------------------------------------
  if (!Array.isArray(src.fields) || src.fields.length === 0) {
    throw new Error('"fields" must be a non-empty array.');
  }
  const fields: IFieldDefinition[] = src.fields.map(normalizeField);

  // ---- list -----------------------------------------------------------
  const legacyTitle = src.listName;
  const listTitle = (src.list && src.list.title) || legacyTitle || '';
  const list = {
    title: listTitle,
    internalName: (src.list && src.list.internalName) || toListInternalName(listTitle),
    description: (src.list && src.list.description) || src.listDescription || ''
  };

  // ---- branding -------------------------------------------------------
  const b = src.branding || {};
  const itemLabel = b.itemLabel || src.itemLabel || 'item';
  const branding = {
    title: b.title || listTitle,
    subtitle: b.subtitle !== undefined ? b.subtitle : (src.listDescription || ''),
    icon: b.icon || src.icon || '🦏',
    accent: b.accent,
    itemLabel,
    itemLabelPlural: b.itemLabelPlural || pluralise(itemLabel)
  };

  // ---- form -----------------------------------------------------------
  const form = normalizeForm(src.form, fields);

  // ---- grid -----------------------------------------------------------
  const grid = normalizeGrid(src.grid, fields);

  return { schemaVersion: 2, list, branding, fields, form, grid };
}

function normalizeField(f: IFieldDefinition): IFieldDefinition {
  if (!f || !f.internalName) throw new Error('Every field needs an "internalName".');
  const type = f.type || 'Text';
  const out: IFieldDefinition = {
    ...f,
    internalName: String(f.internalName),
    displayName: f.displayName || f.internalName,
    type
  };
  if (type === 'DateTime' && out.dateOnly === undefined) out.dateOnly = true;
  if ((type === 'Choice' || type === 'MultiChoice') && !Array.isArray(out.choices)) out.choices = [];
  if (out.internalName.toLowerCase() === 'title') {
    out.internalName = 'Title';
    out.type = 'Text';
    out.required = true;
  }
  // These flags belong to the grid in v2; normaliseGrid reads them then we drop them.
  return out;
}

function normalizeForm(form: Partial<IFormConfig> | undefined, fields: IFieldDefinition[]): IFormConfig {
  const base = buildDefaultForm(fields);
  if (!form) return base;
  const known = fields.map(f => f.internalName).concat(systemFields().map(f => f.internalName));

  let sections: IFormSection[] = Array.isArray(form.sections) && form.sections.length > 0
    ? form.sections.map((s, i) => ({
      id: s.id || `section${i + 1}`,
      title: s.title || '',
      description: s.description,
      columns: s.columns || 2,
      collapsible: !!s.collapsible,
      collapsedByDefault: !!s.collapsedByDefault,
      fields: (s.fields || [])
        .filter(p => p && known.indexOf(p.field) >= 0)
        .map(p => ({ ...p, width: p.width || 'half' }))
    }))
    : base.sections;

  // Any non-system field not placed anywhere gets appended to the last section
  // so a schema change never silently hides a column from the form.
  const placed: Record<string, boolean> = {};
  sections.forEach(s => s.fields.forEach(p => { placed[p.field] = true; }));
  const missing = fields.filter(f => !f.isSystem && !placed[f.internalName]);
  if (missing.length > 0) {
    sections = cloneConfig(sections);
    const last = sections[sections.length - 1];
    missing.forEach(f => last.fields.push({ field: f.internalName, width: f.type === 'Note' ? 'full' : 'half' }));
  }

  return {
    ...base,
    ...form,
    surface: form.surface || base.surface,
    size: form.size || base.size,
    sections
  };
}

function normalizeGrid(grid: Partial<IGridConfig> | undefined, fields: IFieldDefinition[]): IGridConfig {
  const base = buildDefaultGrid(fields);
  if (!grid) return base;
  const byName: Record<string, IFieldDefinition> = {};
  fields.concat(systemFields()).forEach(f => { byName[f.internalName] = f; });

  const columns = Array.isArray(grid.columns) && grid.columns.length > 0
    ? grid.columns
      .filter(c => c && byName[c.field])
      .map(c => ({ ...buildDefaultColumn(byName[c.field]), ...c }))
    : base.columns;

  return { ...base, ...grid, columns };
}

/** Serialise for storage – stable key order, 2-space indent, legacy flags stripped. */
export function serializeConfig(config: IRhinoConfig): string {
  const clean = cloneConfig(config);
  clean.fields = clean.fields.map(f => {
    const { filterable, sortable, ...rest } = f;
    void filterable; void sortable;
    return rest;
  });
  return JSON.stringify(clean, undefined, 2);
}

export function parseConfig(json: string): IRhinoConfig {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch (e) {
    throw new Error(`Not valid JSON: ${(e as Error).message}`);
  }
  return normalizeConfig(raw);
}
