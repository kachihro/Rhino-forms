import { IRhinoConfig } from '../types/RhinoConfig';
import { isValidInternalName, RESERVED_INTERNAL_NAMES, systemFields } from './helpers';

export interface IValidationIssue {
  level: 'error' | 'warning';
  path: string;
  message: string;
}

/**
 * Structural validation of a v2 config. Returns issues instead of throwing so
 * the designer can show them inline. Errors block provisioning / saving;
 * warnings are advisory.
 */
export function validateConfig(config: IRhinoConfig): IValidationIssue[] {
  const issues: IValidationIssue[] = [];
  const err = (path: string, message: string): void => { issues.push({ level: 'error', path, message }); };
  const warn = (path: string, message: string): void => { issues.push({ level: 'warning', path, message }); };

  if (!config.list || !config.list.title || !config.list.title.trim()) {
    err('list.title', 'List title is required.');
  }
  if (config.list && config.list.internalName && !/^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(config.list.internalName)) {
    err('list.internalName', 'List internal name may only contain letters, digits and underscores, and must start with a letter.');
  }

  if (!config.fields || config.fields.length === 0) {
    err('fields', 'At least one field is required.');
    return issues;
  }

  const seen: Record<string, boolean> = {};
  const seenDisplay: Record<string, boolean> = {};
  let hasTitle = false;
  config.fields.forEach((f, i) => {
    const p = `fields[${i}]`;
    if (!f.internalName) { err(p, 'Internal name is required.'); return; }
    if (f.internalName === 'Title') hasTitle = true;
    if (!isValidInternalName(f.internalName)) {
      err(`${p}.internalName`, `"${f.internalName}" is not a valid internal name (letters, digits, underscore; start with a letter; max 32 chars).`);
    }
    if (!f.isSystem && RESERVED_INTERNAL_NAMES.indexOf(f.internalName) >= 0 && f.internalName !== 'Title') {
      err(`${p}.internalName`, `"${f.internalName}" is a reserved SharePoint column name.`);
    }
    const key = f.internalName.toLowerCase();
    if (seen[key]) err(`${p}.internalName`, `Duplicate internal name "${f.internalName}".`);
    seen[key] = true;
    if (!f.displayName || !f.displayName.trim()) err(`${p}.displayName`, 'Display name is required.');
    else {
      const dk = f.displayName.trim().toLowerCase();
      if (seenDisplay[dk]) warn(`${p}.displayName`, `Duplicate display name "${f.displayName}".`);
      seenDisplay[dk] = true;
    }
    if ((f.type === 'Choice' || f.type === 'MultiChoice') && (!f.choices || f.choices.length === 0)) {
      err(`${p}.choices`, `"${f.displayName || f.internalName}" needs at least one choice.`);
    }
    if (f.type === 'Text' && f.maxLength !== undefined && (f.maxLength < 1 || f.maxLength > 255)) {
      err(`${p}.maxLength`, 'Text max length must be between 1 and 255.');
    }
  });
  if (!hasTitle) warn('fields', 'No "Title" field – SharePoint always has one; add it to control its label.');

  const names: Record<string, boolean> = {};
  config.fields.concat(systemFields()).forEach(f => { names[f.internalName] = true; });

  (config.form?.sections || []).forEach((s, si) => {
    s.fields.forEach((p, pi) => {
      if (!names[p.field]) err(`form.sections[${si}].fields[${pi}]`, `Unknown field "${p.field}".`);
    });
  });
  (config.grid?.columns || []).forEach((c, ci) => {
    if (!names[c.field]) err(`grid.columns[${ci}]`, `Unknown field "${c.field}".`);
  });
  if (config.grid?.defaultSort && !names[config.grid.defaultSort.field]) {
    err('grid.defaultSort', `Unknown sort field "${config.grid.defaultSort.field}".`);
  }
  if ((config.grid?.columns || []).length === 0) warn('grid.columns', 'No grid columns selected – the list will look empty.');

  return issues;
}

export function hasErrors(issues: IValidationIssue[]): boolean {
  return issues.some(i => i.level === 'error');
}
