import * as React from 'react';
import {
  TextField, Dropdown, IDropdownOption, Toggle, DatePicker, DayOfWeek, Label, Text, ChoiceGroup
} from '@fluentui/react';
import { IFieldDefinition, IFormFieldPlacement, IUserValue, IUrlValue } from '../../types/RhinoConfig';
import { PeopleService } from '../../services/PeopleService';
import PeoplePicker from '../PeoplePicker/PeoplePicker';
import { formatDate, formatNumber, stringArray, userArray } from '../cells/format';
import { Persona, Badge } from '../cells/CellRenderers';
import styles from './FieldEditor.module.scss';

export interface IFieldEditorProps {
  field: IFieldDefinition;
  placement: IFormFieldPlacement;
  value: unknown;
  onChange: (value: unknown) => void;
  errorMessage?: string;
  readOnly?: boolean;
  peopleService?: PeopleService;
}

const DATE_STRINGS = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Today',
  prevMonthAriaLabel: 'Previous month',
  nextMonthAriaLabel: 'Next month',
  prevYearAriaLabel: 'Previous year',
  nextYearAriaLabel: 'Next year',
  invalidInputErrorMessage: 'Invalid date'
};

function parseAuDate(text: string): Date | undefined {
  const m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/.exec((text || '').trim());
  if (!m) return undefined;
  const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  const d = new Date(year, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? undefined : d;
}

function pad(n: number): string { return n < 10 ? `0${n}` : String(n); }

/** Renders the right input for a field, or a read-only display when readOnly. */
export const FieldEditor: React.FC<IFieldEditorProps> = ({ field, placement, value, onChange, errorMessage, readOnly, peopleService }) => {
  const label = placement.label || field.displayName;
  const help = placement.helpText !== undefined ? placement.helpText : field.description;
  const required = !!field.required && !readOnly;
  const disabled = !!readOnly || !!field.isSystem;

  const helpNode = help ? <Text variant="small" className={styles.help}>{help}</Text> : undefined;

  if (disabled) {
    return (
      <div className={styles.readOnly}>
        <Label>{label}</Label>
        <div className={styles.readOnlyValue}>{renderReadOnly(field, value)}</div>
        {helpNode}
      </div>
    );
  }

  switch (field.type) {
    case 'Text':
      return (
        <div>
          <TextField label={label} required={required} value={(value as string) || ''} placeholder={placement.placeholder}
            onChange={(_, v) => onChange(v)} errorMessage={errorMessage} maxLength={field.maxLength || 255} />
          {helpNode}
        </div>
      );
    case 'Note':
      return (
        <div>
          <TextField label={label} required={required} multiline autoAdjustHeight rows={4} value={(value as string) || ''}
            placeholder={placement.placeholder} onChange={(_, v) => onChange(v)} errorMessage={errorMessage} />
          {helpNode}
        </div>
      );
    case 'Choice': {
      const choices = field.choices || [];
      if (choices.length > 0 && choices.length <= 4 && !placement.placeholder) {
        return (
          <div>
            <ChoiceGroup label={label} required={required} selectedKey={(value as string) || undefined}
              options={choices.map(c => ({ key: c, text: c }))}
              onChange={(_, opt) => onChange(opt?.key)} className={styles.inlineChoice} />
            {errorMessage && <Text variant="small" className={styles.error}>{errorMessage}</Text>}
            {helpNode}
          </div>
        );
      }
      const options: IDropdownOption[] = choices.map(c => ({ key: c, text: c }));
      return (
        <div>
          <Dropdown label={label} required={required} options={options} selectedKey={(value as string) || null}
            placeholder={placement.placeholder || 'Select…'} onChange={(_, opt) => onChange(opt?.key)} errorMessage={errorMessage} />
          {helpNode}
        </div>
      );
    }
    case 'MultiChoice': {
      const options: IDropdownOption[] = (field.choices || []).map(c => ({ key: c, text: c }));
      const selected = stringArray(value);
      return (
        <div>
          <Dropdown label={label} required={required} multiSelect options={options} selectedKeys={selected}
            placeholder={placement.placeholder || 'Select one or more…'}
            onChange={(_, opt) => {
              if (!opt) return;
              const next = opt.selected ? selected.concat([String(opt.key)]) : selected.filter(s => s !== opt.key);
              onChange(next);
            }}
            errorMessage={errorMessage} />
          {helpNode}
        </div>
      );
    }
    case 'DateTime': {
      const d = value instanceof Date && !isNaN(value.getTime()) ? value : undefined;
      return (
        <div>
          <DatePicker label={label} isRequired={required} value={d} placeholder={placement.placeholder || 'dd/mm/yyyy'}
            firstDayOfWeek={DayOfWeek.Monday} strings={DATE_STRINGS} allowTextInput
            formatDate={dt => dt ? formatDate(dt) : ''} parseDateFromString={parseAuDate}
            onSelectDate={dt => onChange(dt || undefined)} />
          {field.dateOnly === false && (
            <TextField type="time" className={styles.timeField} value={d ? `${pad(d.getHours())}:${pad(d.getMinutes())}` : ''}
              disabled={!d}
              onChange={(_, v) => {
                if (!d || !v) return;
                const parts = v.split(':').map(Number);
                const next = new Date(d.getTime());
                next.setHours(parts[0] || 0, parts[1] || 0, 0, 0);
                onChange(next);
              }} />
          )}
          {errorMessage && <Text variant="small" className={styles.error}>{errorMessage}</Text>}
          {helpNode}
        </div>
      );
    }
    case 'Number': {
      const isInt = field.numberFormat === 'integer';
      const prefix = field.numberFormat === 'currency' ? '$' : undefined;
      const suffix = field.numberFormat === 'percent' ? '%' : undefined;
      const step = isInt ? '1' : (field.decimals !== undefined ? String(Math.pow(10, -field.decimals)) : 'any');
      return (
        <div>
          <TextField label={label} required={required} type="number" prefix={prefix} suffix={suffix} step={step as any}
            value={value === undefined || value === null ? '' : String(value)} placeholder={placement.placeholder}
            onChange={(_, v) => onChange(v === '' || v === undefined ? undefined : Number(v))} errorMessage={errorMessage} />
          {helpNode}
        </div>
      );
    }
    case 'Boolean':
      return (
        <div className={styles.toggleWrap}>
          <Toggle label={label} checked={!!value} onText="Yes" offText="No" onChange={(_, c) => onChange(!!c)} />
          {helpNode}
        </div>
      );
    case 'User':
      return (
        <div>
          <PeoplePicker label={label} required={required} multiple={!!field.allowMultiple} value={userArray(value)}
            onChange={(users: IUserValue[]) => onChange(users)} peopleService={peopleService} errorMessage={errorMessage}
            placeholder={placement.placeholder} />
          {helpNode}
        </div>
      );
    case 'URL': {
      const u = (value || {}) as IUrlValue;
      return (
        <div>
          <TextField label={label} required={required} value={u.Url || ''} placeholder={placement.placeholder || 'https://…'}
            onChange={(_, v) => onChange({ ...u, Url: v })} errorMessage={errorMessage} iconProps={{ iconName: 'Link' }} />
          <TextField className={styles.urlDesc} value={u.Description || ''} placeholder="Link text (optional)"
            onChange={(_, v) => onChange({ ...u, Description: v })} />
          {helpNode}
        </div>
      );
    }
    default:
      return <TextField label={label} value={String(value || '')} onChange={(_, v) => onChange(v)} />;
  }
};

export function renderReadOnly(field: IFieldDefinition, value: unknown): React.ReactNode {
  if (value === null || value === undefined || value === '') return <span className={styles.empty}>—</span>;
  switch (field.type) {
    case 'User': {
      const users = userArray(value);
      if (users.length === 0) return <span className={styles.empty}>—</span>;
      return <div className={styles.personaList}>{users.map(u => <Persona key={u.Id} name={u.Title || ''} email={u.EMail} />)}</div>;
    }
    case 'Choice': return <Badge value={String(value)} />;
    case 'MultiChoice': {
      const arr = stringArray(value);
      return arr.length ? <div className={styles.tagList}>{arr.map(v => <Badge key={v} value={v} />)}</div> : <span className={styles.empty}>—</span>;
    }
    case 'DateTime': return formatDate(value, field.dateOnly === false);
    case 'Number': return formatNumber(value, field);
    case 'Boolean': return value ? 'Yes' : 'No';
    case 'URL': {
      const u = value as IUrlValue;
      return u.Url ? <a href={u.Url} target="_blank" rel="noopener noreferrer">{u.Description || u.Url}</a> : <span className={styles.empty}>—</span>;
    }
    case 'Note': return <div className={styles.multiline}>{String(value)}</div>;
    default: return String(value);
  }
}

export default FieldEditor;
