import { IFieldDefinition, IUserValue } from '../../types/RhinoConfig';

const LOCALE = 'en-AU';

export function formatDate(value: unknown, withTime: boolean = false): string {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(String(value));
  if (isNaN(d.getTime())) return String(value);
  const date = d.toLocaleDateString(LOCALE, { day: '2-digit', month: '2-digit', year: 'numeric' });
  if (!withTime) return date;
  return `${date} ${d.toLocaleTimeString(LOCALE, { hour: '2-digit', minute: '2-digit' })}`;
}

export function formatNumber(value: unknown, field?: IFieldDefinition): string {
  if (value === null || value === undefined || value === '') return '';
  const n = typeof value === 'number' ? value : Number(value);
  if (isNaN(n)) return String(value);
  const fmt = field?.numberFormat || 'decimal';
  const decimals = field?.decimals;
  switch (fmt) {
    case 'currency':
      return n.toLocaleString(LOCALE, { style: 'currency', currency: 'AUD', minimumFractionDigits: decimals ?? 2, maximumFractionDigits: decimals ?? 2 });
    case 'percent':
      return `${n.toLocaleString(LOCALE, { maximumFractionDigits: decimals ?? 1 })}%`;
    case 'integer':
      return Math.round(n).toLocaleString(LOCALE);
    default:
      return n.toLocaleString(LOCALE, { maximumFractionDigits: decimals ?? 2 });
  }
}

export function userArray(value: unknown): IUserValue[] {
  if (!value) return [];
  if (Array.isArray(value)) return value.filter(Boolean) as IUserValue[];
  return [value as IUserValue];
}

export function stringArray(value: unknown): string[] {
  if (!value) return [];
  if (Array.isArray(value)) return value.map(String);
  return [String(value)];
}

/** Plain-text representation of any cell value (used for search + CSV export). */
export function toPlainText(value: unknown, field: IFieldDefinition): string {
  if (value === null || value === undefined) return '';
  switch (field.type) {
    case 'User': return userArray(value).map(u => u.Title || '').join(', ');
    case 'MultiChoice': return stringArray(value).join(', ');
    case 'DateTime': return formatDate(value, field.dateOnly === false);
    case 'Number': return formatNumber(value, field);
    case 'Boolean': return value ? 'Yes' : 'No';
    case 'URL': return (value as { Url?: string })?.Url || '';
    default: return String(value);
  }
}

export function getInitials(name: string): string {
  const parts = (name || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  return parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : parts[0].slice(0, 2).toUpperCase();
}

const AVATAR_COLORS = ['#0f6cbd', '#107c10', '#8e3ba3', '#ca5010', '#038387', '#7160e8', '#004b50', '#c239b3'];

export function getAvatarColor(name: string): string {
  let h = 5381;
  for (let i = 0; i < (name || '').length; i++) h = ((h << 5) + h) ^ name.charCodeAt(i);
  return AVATAR_COLORS[Math.abs(h) % AVATAR_COLORS.length];
}
