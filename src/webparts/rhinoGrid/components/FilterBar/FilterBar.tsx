import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Callout,
  DirectionalHint,
  SearchBox,
  Dropdown,
  IDropdownOption,
  DatePicker,
  TextField,
  DefaultButton,
  IconButton,
  Text
} from '@fluentui/react';
import { IFieldDefinition, IFilterState, IFilterValue } from '../../types/IFieldConfig';
import styles from './FilterBar.module.scss';

export interface IFilterBarProps {
  fields: IFieldDefinition[];
  filterState: IFilterState;
  onFilterChange: (newState: IFilterState) => void;
}

interface IOpenChip {
  fieldName: string;
  target: HTMLElement;
}

export const FilterBar: React.FC<IFilterBarProps> = ({ fields, filterState, onFilterChange }) => {
  const [localState, setLocalState] = useState<IFilterState>(filterState);
  const [openChip, setOpenChip] = useState<IOpenChip | null>(null);
  const debounceTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const filterableFields = fields.filter(f => f.filterable);

  useEffect(() => { setLocalState(filterState); }, [filterState]);

  if (filterableFields.length === 0) return null;

  const updateFilter = (fieldName: string, value: IFilterValue): void => {
    const newState: IFilterState = { ...localState, [fieldName]: value };
    setLocalState(newState);
    if (debounceTimer.current) clearTimeout(debounceTimer.current);
    debounceTimer.current = setTimeout(() => onFilterChange(newState), 300);
  };

  const clearField = (fieldName: string): void => {
    const newState: IFilterState = { ...localState };
    delete newState[fieldName];
    setLocalState(newState);
    onFilterChange(newState);
  };

  const clearAll = (): void => {
    setLocalState({});
    onFilterChange({});
  };

  const hasAnyFilter = (field: IFieldDefinition): boolean => {
    const v = localState[field.internalName];
    if (!v) return false;
    return !!(v.text || v.choice || v.dateFrom || v.dateTo ||
      v.numberMin !== undefined || v.numberMax !== undefined ||
      v.boolValue !== undefined);
  };

  const getChipLabel = (field: IFieldDefinition): string => {
    const v = localState[field.internalName];
    if (!v) return field.displayName;
    if (v.text) return `${field.displayName}: ${v.text}`;
    if (v.choice) return `${field.displayName}: ${v.choice}`;
    if (v.boolValue !== undefined) return `${field.displayName}: ${v.boolValue ? 'Yes' : 'No'}`;
    if (v.dateFrom && v.dateTo) return `${field.displayName}: ${v.dateFrom} – ${v.dateTo}`;
    if (v.dateFrom) return `${field.displayName}: from ${v.dateFrom}`;
    if (v.dateTo) return `${field.displayName}: to ${v.dateTo}`;
    if (v.numberMin !== undefined && v.numberMax !== undefined) return `${field.displayName}: ${v.numberMin}–${v.numberMax}`;
    if (v.numberMin !== undefined) return `${field.displayName}: ≥${v.numberMin}`;
    if (v.numberMax !== undefined) return `${field.displayName}: ≤${v.numberMax}`;
    return field.displayName;
  };

  const renderCalloutContent = (field: IFieldDefinition): JSX.Element => {
    const current = localState[field.internalName] || {};
    switch (field.type) {
      case 'Text':
      case 'Note':
        return (
          <div className={styles.calloutBody}>
            <SearchBox
              placeholder={`Search ${field.displayName}`}
              value={current.text || ''}
              onChange={(_, v) => updateFilter(field.internalName, { ...current, text: v || '' })}
              onClear={() => updateFilter(field.internalName, { ...current, text: '' })}
              autoFocus
            />
          </div>
        );
      case 'Choice': {
        const options: IDropdownOption[] = [
          { key: '', text: `All` },
          ...(field.choices || []).map(c => ({ key: c, text: c }))
        ];
        return (
          <div className={styles.calloutBody}>
            <Dropdown
              label={field.displayName}
              options={options}
              selectedKey={current.choice || ''}
              onChange={(_, opt) => updateFilter(field.internalName, { ...current, choice: (opt?.key as string) || '' })}
            />
          </div>
        );
      }
      case 'DateTime':
        return (
          <div className={styles.calloutBody}>
            <DatePicker
              label="From"
              placeholder="Start date"
              value={current.dateFrom ? new Date(current.dateFrom) : undefined}
              onSelectDate={d => updateFilter(field.internalName, { ...current, dateFrom: d ? d.toISOString().split('T')[0] : undefined })}
            />
            <DatePicker
              label="To"
              placeholder="End date"
              value={current.dateTo ? new Date(current.dateTo) : undefined}
              onSelectDate={d => updateFilter(field.internalName, { ...current, dateTo: d ? d.toISOString().split('T')[0] : undefined })}
            />
          </div>
        );
      case 'Number':
        return (
          <div className={styles.calloutBody}>
            <TextField
              label="Min"
              type="number"
              value={current.numberMin !== undefined ? String(current.numberMin) : ''}
              onChange={(_, v) => updateFilter(field.internalName, { ...current, numberMin: v ? Number(v) : undefined })}
            />
            <TextField
              label="Max"
              type="number"
              value={current.numberMax !== undefined ? String(current.numberMax) : ''}
              onChange={(_, v) => updateFilter(field.internalName, { ...current, numberMax: v ? Number(v) : undefined })}
            />
          </div>
        );
      case 'Boolean':
        return (
          <div className={styles.calloutBody}>
            <Dropdown
              label={field.displayName}
              options={[
                { key: '', text: 'Any' },
                { key: 'true', text: 'Yes' },
                { key: 'false', text: 'No' }
              ]}
              selectedKey={current.boolValue === true ? 'true' : current.boolValue === false ? 'false' : ''}
              onChange={(_, opt) => {
                const bv = opt?.key === 'true' ? true : opt?.key === 'false' ? false : undefined;
                updateFilter(field.internalName, { ...current, boolValue: bv });
              }}
            />
          </div>
        );
      default:
        return <></>;
    }
  };

  const hasFilters = filterableFields.some(f => hasAnyFilter(f));

  return (
    <div className={styles.filterBar}>
      <Text className={styles.filterLabel}>FILTER</Text>
      {filterableFields.map(field => {
        const active = hasAnyFilter(field);
        return (
          <div key={field.internalName} className={`${styles.chip} ${active ? styles.chipActive : ''}`}>
            <button
              className={styles.chipBtn}
              onClick={(e) => setOpenChip(openChip?.fieldName === field.internalName ? null : { fieldName: field.internalName, target: e.currentTarget as HTMLElement })}
            >
              <span className={styles.chipLabel}>{getChipLabel(field)}</span>
              <span className={styles.chipCaret}>▾</span>
            </button>
            {active && (
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                className={styles.chipClear}
                title="Clear"
                onClick={() => clearField(field.internalName)}
              />
            )}
          </div>
        );
      })}

      {hasFilters && (
        <DefaultButton
          text="Clear all"
          iconProps={{ iconName: 'ClearFilter' }}
          className={styles.clearAllBtn}
          onClick={clearAll}
        />
      )}

      {openChip && (
        <Callout
          target={openChip.target}
          onDismiss={() => setOpenChip(null)}
          directionalHint={DirectionalHint.bottomLeftEdge}
          calloutMinWidth={220}
          calloutMaxWidth={320}
          isBeakVisible={false}
          styles={{ root: { boxShadow: '0 4px 16px rgba(0,0,0,0.12)', borderRadius: 6 } }}
        >
          {renderCalloutContent(filterableFields.find(f => f.internalName === openChip.fieldName)!)}
        </Callout>
      )}
    </div>
  );
};

export default FilterBar;
