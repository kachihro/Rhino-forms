import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Callout, DirectionalHint, SearchBox, Dropdown, IDropdownOption, DatePicker, DayOfWeek, TextField, DefaultButton, IconButton, Text
} from '@fluentui/react';
import { IFieldDefinition, IFilterState, IFilterValue, IGridColumn } from '../../types/RhinoConfig';
import { formatDate } from '../cells/format';
import styles from './FilterBar.module.scss';

export interface IFilterBarProps {
  fields: IFieldDefinition[];
  columns: IGridColumn[];
  filterState: IFilterState;
  onFilterChange: (newState: IFilterState) => void;
}

interface IOpenChip {
  fieldName: string;
  target: HTMLElement;
}

function isoDate(d: Date): string {
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function hasValue(v?: IFilterValue): boolean {
  if (!v) return false;
  return !!(v.text || v.choice || v.dateFrom || v.dateTo || v.userTitle ||
    v.numberMin !== undefined || v.numberMax !== undefined || v.boolValue !== undefined);
}

export const FilterBar: React.FC<IFilterBarProps> = ({ fields, columns, filterState, onFilterChange }) => {
  const [localState, setLocalState] = useState<IFilterState>(filterState);
  const [openChip, setOpenChip] = useState<IOpenChip | undefined>();
  const debounce = useRef<ReturnType<typeof setTimeout> | undefined>();

  useEffect(() => { setLocalState(filterState); }, [filterState]);

  const filterable = columns
    .filter(c => c.filterable)
    .map(c => ({ column: c, field: fields.filter(f => f.internalName === c.field)[0] }))
    .filter(x => !!x.field && x.field.type !== 'URL');

  if (filterable.length === 0) return <div />;

  const update = (name: string, value: IFilterValue, immediate: boolean = false): void => {
    const next: IFilterState = { ...localState, [name]: value };
    setLocalState(next);
    if (debounce.current) clearTimeout(debounce.current);
    if (immediate) onFilterChange(next);
    else debounce.current = setTimeout(() => onFilterChange(next), 350);
  };

  const clearField = (name: string): void => {
    const next: IFilterState = { ...localState };
    delete next[name];
    setLocalState(next);
    onFilterChange(next);
  };

  const clearAll = (): void => {
    setLocalState({});
    onFilterChange({});
  };

  const chipLabel = (field: IFieldDefinition, column: IGridColumn): string => {
    const label = column.label || field.displayName;
    const v = localState[field.internalName];
    if (!v) return label;
    if (v.text) return `${label}: ${v.text}`;
    if (v.choice) return `${label}: ${v.choice}`;
    if (v.userTitle) return `${label}: ${v.userTitle}`;
    if (v.boolValue !== undefined) return `${label}: ${v.boolValue ? 'Yes' : 'No'}`;
    const fd = v.dateFrom ? formatDate(v.dateFrom) : '';
    const td = v.dateTo ? formatDate(v.dateTo) : '';
    if (fd && td) return `${label}: ${fd} – ${td}`;
    if (fd) return `${label}: from ${fd}`;
    if (td) return `${label}: to ${td}`;
    if (v.numberMin !== undefined && v.numberMax !== undefined) return `${label}: ${v.numberMin}–${v.numberMax}`;
    if (v.numberMin !== undefined) return `${label}: ≥ ${v.numberMin}`;
    if (v.numberMax !== undefined) return `${label}: ≤ ${v.numberMax}`;
    return label;
  };

  const renderCallout = (field: IFieldDefinition): JSX.Element => {
    const cur = localState[field.internalName] || {};
    const n = field.internalName;
    switch (field.type) {
      case 'Text':
      case 'Note':
        return (
          <div className={styles.calloutBody}>
            <SearchBox placeholder={`Contains…`} value={cur.text || ''} autoFocus
              onChange={(_, v) => update(n, { ...cur, text: v || '' })}
              onClear={() => update(n, { ...cur, text: '' }, true)} />
          </div>
        );
      case 'User':
        return (
          <div className={styles.calloutBody}>
            <SearchBox placeholder="Name contains…" value={cur.userTitle || ''} autoFocus
              onChange={(_, v) => update(n, { ...cur, userTitle: v || '' })}
              onClear={() => update(n, { ...cur, userTitle: '' }, true)} />
          </div>
        );
      case 'Choice':
      case 'MultiChoice': {
        const options: IDropdownOption[] = [{ key: '', text: 'All' }].concat((field.choices || []).map(c => ({ key: c, text: c })));
        return (
          <div className={styles.calloutBody}>
            <Dropdown options={options} selectedKey={cur.choice || ''}
              onChange={(_, opt) => update(n, { ...cur, choice: (opt?.key as string) || '' }, true)} />
          </div>
        );
      }
      case 'DateTime':
        return (
          <div className={styles.calloutBody}>
            <DatePicker label="From" placeholder="dd/mm/yyyy" firstDayOfWeek={DayOfWeek.Monday} formatDate={d => d ? formatDate(d) : ''}
              value={cur.dateFrom ? new Date(cur.dateFrom) : undefined}
              onSelectDate={d => update(n, { ...cur, dateFrom: d ? isoDate(d) : undefined }, true)} />
            <DatePicker label="To" placeholder="dd/mm/yyyy" firstDayOfWeek={DayOfWeek.Monday} formatDate={d => d ? formatDate(d) : ''}
              value={cur.dateTo ? new Date(cur.dateTo) : undefined}
              onSelectDate={d => update(n, { ...cur, dateTo: d ? isoDate(d) : undefined }, true)} />
          </div>
        );
      case 'Number':
        return (
          <div className={styles.calloutBody}>
            <TextField label="Min" type="number" value={cur.numberMin !== undefined ? String(cur.numberMin) : ''}
              onChange={(_, v) => update(n, { ...cur, numberMin: v ? Number(v) : undefined })} />
            <TextField label="Max" type="number" value={cur.numberMax !== undefined ? String(cur.numberMax) : ''}
              onChange={(_, v) => update(n, { ...cur, numberMax: v ? Number(v) : undefined })} />
          </div>
        );
      case 'Boolean':
        return (
          <div className={styles.calloutBody}>
            <Dropdown options={[{ key: '', text: 'Any' }, { key: 'true', text: 'Yes' }, { key: 'false', text: 'No' }]}
              selectedKey={cur.boolValue === true ? 'true' : cur.boolValue === false ? 'false' : ''}
              onChange={(_, opt) => update(n, { ...cur, boolValue: opt?.key === 'true' ? true : opt?.key === 'false' ? false : undefined }, true)} />
          </div>
        );
      default:
        return <div />;
    }
  };

  const anyActive = filterable.some(x => hasValue(localState[x.field.internalName]));
  const open = openChip ? filterable.filter(x => x.field.internalName === openChip.fieldName)[0] : undefined;

  return (
    <div className={styles.filterBar}>
      <Text className={styles.filterLabel}>Filter</Text>
      {filterable.map(({ field, column }) => {
        const active = hasValue(localState[field.internalName]);
        return (
          <div key={field.internalName} className={`${styles.chip} ${active ? styles.chipActive : ''}`}>
            <button
              type="button"
              className={styles.chipBtn}
              onClick={e => setOpenChip(openChip?.fieldName === field.internalName ? undefined : { fieldName: field.internalName, target: e.currentTarget as HTMLElement })}
            >
              <span className={styles.chipText}>{chipLabel(field, column)}</span>
              <span className={styles.chipCaret}>▾</span>
            </button>
            {active && (
              <IconButton iconProps={{ iconName: 'Cancel' }} className={styles.chipClear} title="Clear" onClick={() => clearField(field.internalName)} />
            )}
          </div>
        );
      })}
      {anyActive && (
        <DefaultButton text="Clear all" iconProps={{ iconName: 'ClearFilter' }} className={styles.clearAllBtn} onClick={clearAll} />
      )}
      {openChip && open && (
        <Callout
          target={openChip.target}
          onDismiss={() => setOpenChip(undefined)}
          directionalHint={DirectionalHint.bottomLeftEdge}
          calloutMinWidth={230}
          calloutMaxWidth={320}
          isBeakVisible={false}
          styles={{ root: { boxShadow: '0 6px 20px rgba(0,0,0,0.14)', borderRadius: 8 } }}
        >
          <div className={styles.calloutTitle}>{open.column.label || open.field.displayName}</div>
          {renderCallout(open.field)}
        </Callout>
      )}
    </div>
  );
};

export default FilterBar;
