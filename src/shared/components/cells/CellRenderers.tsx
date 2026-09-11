import * as React from 'react';
import { Icon, Link } from '@fluentui/react';
import { IFieldDefinition, IGridColumn, IGridItem, BadgeTone, CellRenderer, IUrlValue } from '../../types/RhinoConfig';
import { defaultRenderer } from '../../config/helpers';
import { formatDate, formatNumber, userArray, stringArray, getInitials, getAvatarColor } from './format';
import styles from './Cells.module.scss';

export const TONES: Record<BadgeTone, { bg: string; fg: string; dot: string }> = {
  neutral: { bg: '#f0f0f0', fg: '#424242', dot: '#7a7a7a' },
  info: { bg: '#e6f0fb', fg: '#0f4c8a', dot: '#0f6cbd' },
  success: { bg: '#e3f5e3', fg: '#0e6b0e', dot: '#107c10' },
  warning: { bg: '#fdf3d7', fg: '#7a4b00', dot: '#c19c00' },
  danger: { bg: '#fde7e9', fg: '#a4262c', dot: '#c50f1f' },
  purple: { bg: '#efe9fb', fg: '#5b2ea3', dot: '#7160e8' },
  teal: { bg: '#e0f4f4', fg: '#005f60', dot: '#038387' }
};

export function autoTone(value: string): BadgeTone {
  const v = (value || '').toLowerCase();
  if (/pend|progress|wait|review|submitted|open/.test(v)) return 'warning';
  if (/approv|complet|active|done|success|paid|closed won|yes/.test(v)) return 'success';
  if (/reject|cancel|fail|deni|overdue|blocked|critical|high|urgent/.test(v)) return 'danger';
  if (/hold|paused|defer|suspend|archiv|not started|low/.test(v)) return 'neutral';
  if (/medium|normal/.test(v)) return 'info';
  return 'info';
}

export function toneFor(value: string, column?: IGridColumn): BadgeTone {
  const explicit = column?.badgeTones && column.badgeTones[value];
  return explicit || autoTone(value);
}

export const Badge: React.FC<{ value: string; tone?: BadgeTone }> = ({ value, tone }) => {
  const t = TONES[tone || autoTone(value)];
  return (
    <span className={styles.badge} style={{ background: t.bg, color: t.fg }}>
      <span className={styles.badgeDot} style={{ background: t.dot }} />
      {value}
    </span>
  );
};

export const Persona: React.FC<{ name: string; email?: string; compact?: boolean }> = ({ name, email, compact }) => (
  <span className={styles.persona} title={email || name}>
    <span className={styles.avatar} style={{ background: getAvatarColor(name) }}>{getInitials(name)}</span>
    {!compact && <span className={styles.personaName}>{name}</span>}
  </span>
);

export function resolveRenderer(field: IFieldDefinition, column?: IGridColumn): CellRenderer {
  const r = column?.render;
  if (!r || r === 'auto') return defaultRenderer(field.type);
  return r;
}

export function renderCell(item: IGridItem, field: IFieldDefinition, column?: IGridColumn): React.ReactNode {
  const value = item[field.internalName];
  if (value === null || value === undefined || value === '') return '';
  const renderer = resolveRenderer(field, column);

  switch (renderer) {
    case 'badge':
      return <Badge value={String(value)} tone={toneFor(String(value), column)} />;
    case 'tags':
      return (
        <span className={styles.tags}>
          {stringArray(value).map(v => <Badge key={v} value={v} tone={toneFor(v, column)} />)}
        </span>
      );
    case 'persona': {
      const users = userArray(value);
      if (users.length === 0) return '';
      if (users.length === 1) return <Persona name={users[0].Title || ''} email={users[0].EMail} />;
      return (
        <span className={styles.personaStack}>
          {users.slice(0, 3).map(u => <Persona key={u.Id} name={u.Title || ''} email={u.EMail} compact />)}
          <span className={styles.personaMore}>{users.length > 3 ? `+${users.length - 3}` : users.map(u => u.Title).join(', ')}</span>
        </span>
      );
    }
    case 'date':
      return <span className={styles.mono}>{formatDate(value, false)}</span>;
    case 'datetime':
      return <span className={styles.mono}>{formatDate(value, true)}</span>;
    case 'number':
      return <span className={styles.numeric}>{formatNumber(value, field)}</span>;
    case 'currency':
      return <span className={styles.numeric}>{formatNumber(value, { ...field, numberFormat: 'currency' })}</span>;
    case 'percent':
      return <span className={styles.numeric}>{formatNumber(value, { ...field, numberFormat: 'percent' })}</span>;
    case 'check':
      return value
        ? <Icon iconName="SkypeCircleCheck" className={styles.boolTrue} />
        : <Icon iconName="StatusCircleRing" className={styles.boolFalse} />;
    case 'link': {
      const u = value as IUrlValue;
      return u?.Url ? <Link href={u.Url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()}>{u.Description || u.Url}</Link> : '';
    }
    case 'text':
    default:
      if (field.type === 'Note') return <span className={styles.noteCell}>{String(value)}</span>;
      if (field.type === 'DateTime') return formatDate(value, field.dateOnly === false);
      if (field.type === 'Number') return formatNumber(value, field);
      if (field.type === 'Boolean') return value ? 'Yes' : 'No';
      if (field.type === 'User') return userArray(value).map(u => u.Title).join(', ');
      if (field.type === 'MultiChoice') return stringArray(value).join(', ');
      return String(value);
  }
}
