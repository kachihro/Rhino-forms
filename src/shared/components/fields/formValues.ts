import { IFieldDefinition, IGridItem, IUserValue, IUrlValue } from '../../types/RhinoConfig';
import { userArray, stringArray } from '../cells/format';

/**
 * Canonical in-form value per field type:
 *   Text / Note / Choice -> string
 *   MultiChoice          -> string[]
 *   DateTime             -> Date | undefined
 *   Number               -> number | undefined
 *   Boolean              -> boolean
 *   User                 -> IUserValue[]
 *   URL                  -> IUrlValue
 */
export type FormValues = Record<string, unknown>;

export function defaultValues(fields: IFieldDefinition[]): FormValues {
  const out: FormValues = {};
  fields.forEach(f => {
    if (f.isSystem) return;
    const d = f.defaultValue;
    switch (f.type) {
      case 'MultiChoice': out[f.internalName] = d === undefined ? [] : stringArray(d); break;
      case 'Boolean': out[f.internalName] = !!d; break;
      case 'User': out[f.internalName] = []; break;
      case 'URL': out[f.internalName] = { Url: '', Description: '' }; break;
      case 'DateTime': out[f.internalName] = d === 'today' ? new Date() : (d ? new Date(String(d)) : undefined); break;
      case 'Number': out[f.internalName] = d === undefined || d === '' ? undefined : Number(d); break;
      default: out[f.internalName] = d === undefined ? '' : String(d);
    }
  });
  return out;
}

export function itemToValues(item: IGridItem, fields: IFieldDefinition[]): FormValues {
  const out: FormValues = {};
  fields.forEach(f => {
    const raw = item[f.internalName];
    switch (f.type) {
      case 'MultiChoice': out[f.internalName] = stringArray(raw); break;
      case 'Boolean': out[f.internalName] = !!raw; break;
      case 'User': out[f.internalName] = userArray(raw); break;
      case 'URL': out[f.internalName] = raw ? { Url: (raw as IUrlValue).Url || '', Description: (raw as IUrlValue).Description || '' } : { Url: '', Description: '' }; break;
      case 'DateTime': out[f.internalName] = raw ? new Date(String(raw)) : undefined; break;
      case 'Number': out[f.internalName] = raw === null || raw === undefined ? undefined : Number(raw); break;
      default: out[f.internalName] = raw === null || raw === undefined ? '' : String(raw);
    }
  });
  return out;
}

/** Build the REST payload. Only editable, non-system fields are written. */
export function valuesToPayload(values: FormValues, fields: IFieldDefinition[], editable: (name: string) => boolean): Record<string, unknown> {
  const payload: Record<string, unknown> = {};
  fields.forEach(f => {
    if (f.isSystem || !editable(f.internalName)) return;
    const v = values[f.internalName];
    switch (f.type) {
      case 'MultiChoice':
        payload[f.internalName] = stringArray(v);
        break;
      case 'User': {
        const ids = userArray(v).map(u => u.Id);
        payload[`${f.internalName}Id`] = f.allowMultiple ? ids : (ids.length ? ids[0] : null);
        break;
      }
      case 'URL': {
        const u = (v || {}) as IUrlValue;
        payload[f.internalName] = u.Url ? { Url: u.Url, Description: u.Description || u.Url } : null;
        break;
      }
      case 'DateTime': {
        const d = v as Date | undefined;
        if (!d || isNaN(d.getTime())) { payload[f.internalName] = null; break; }
        payload[f.internalName] = f.dateOnly === false ? d.toISOString() : toLocalNoonIso(d);
        break;
      }
      case 'Number':
        payload[f.internalName] = v === undefined || v === '' || v === null ? null : Number(v);
        break;
      case 'Boolean':
        payload[f.internalName] = !!v;
        break;
      default:
        payload[f.internalName] = v === undefined || v === null ? '' : String(v);
    }
  });
  return payload;
}

/** Date-only columns: send local midday so timezone shifts never move the calendar day. */
function toLocalNoonIso(d: Date): string {
  const local = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
  return local.toISOString();
}

export function isEmptyValue(v: unknown, field: IFieldDefinition): boolean {
  if (v === undefined || v === null) return true;
  switch (field.type) {
    case 'MultiChoice':
    case 'User': return (v as unknown[]).length === 0;
    case 'URL': return !(v as IUrlValue).Url;
    case 'Boolean': return false;
    case 'DateTime': return !(v instanceof Date) || isNaN(v.getTime());
    case 'Number': return v === '' || (typeof v === 'number' && isNaN(v));
    default: return String(v).trim() === '';
  }
}

export function validateValues(values: FormValues, fields: IFieldDefinition[], editable: (name: string) => boolean): Record<string, string> {
  const errors: Record<string, string> = {};
  fields.forEach(f => {
    if (f.isSystem || !editable(f.internalName)) return;
    const v = values[f.internalName];
    if (f.required && isEmptyValue(v, f)) {
      errors[f.internalName] = `${f.displayName} is required`;
      return;
    }
    if (f.type === 'Number' && !isEmptyValue(v, f)) {
      const n = Number(v);
      if (f.min !== undefined && n < f.min) errors[f.internalName] = `Must be at least ${f.min}`;
      if (f.max !== undefined && n > f.max) errors[f.internalName] = `Must be at most ${f.max}`;
      if (f.numberFormat === 'integer' && n % 1 !== 0) errors[f.internalName] = 'Must be a whole number';
    }
    if (f.type === 'Text' && f.maxLength && String(v || '').length > f.maxLength) {
      errors[f.internalName] = `Maximum ${f.maxLength} characters`;
    }
    if (f.type === 'URL' && !isEmptyValue(v, f) && !/^https?:\/\//i.test((v as IUrlValue).Url || '')) {
      errors[f.internalName] = 'Enter a full URL starting with http:// or https://';
    }
  });
  return errors;
}

export function toUserValues(v: unknown): IUserValue[] { return userArray(v); }
