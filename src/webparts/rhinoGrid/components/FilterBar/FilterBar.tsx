import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  SearchBox,
  Dropdown,
  IDropdownOption,
  DatePicker,
  TextField,
  DefaultButton,
  Label
} from '@fluentui/react';
import { IFieldDefinition, IFilterState, IFilterValue } from '../../types/IFieldConfig';
import styles from './FilterBar.module.scss';

export interface IFilterBarProps {
  fields: IFieldDefinition[];
  filterState: IFilterState;
  onFilterChange: (newState: IFilterState) => void;
}

export const FilterBar: React.FC<IFilterBarProps> = ({ fields, filterState, onFilterChange }) => {
  const [localState, setLocalState] = useState<IFilterState>(filterState);
  const debounceTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const filterableFields = fields.filter(f => f.filterable);

  useEffect(() => {
    setLocalState(filterState);
  }, [filterState]);

  const updateFilter = (fieldName: string, value: IFilterValue): void => {
    const newState: IFilterState = { ...localState, [fieldName]: value };
    setLocalState(newState);
    if (debounceTimer.current) clearTimeout(debounceTimer.current);
    debounceTimer.current = setTimeout(() => {
      onFilterChange(newState);
    }, 300);
  };

  const clearAll = (): void => {
    const empty: IFilterState = {};
    setLocalState(empty);
    onFilterChange(empty);
  };

  const hasFilters = Object.keys(localState).some(k => {
    const v = localState[k];
    return v && (v.text || v.choice || v.dateFrom || v.dateTo || v.numberMin !== undefined || v.numberMax !== undefined || v.boolValue !== null && v.boolValue !== undefined);
  });

  if (filterableFields.length === 0) return null;

  const renderFilterControl = (field: IFieldDefinition): JSX.Element => {
    const current = localState[field.internalName] || {};

    switch (field.type) {
      case 'Text':
      case 'Note':
        return (
          <div key={field.internalName} className={styles.filterField}>
            <SearchBox
              placeholder={`Filter by ${field.displayName}`}
              value={current.text || ''}
              onChange={(_, val) => updateFilter(field.internalName, { ...current, text: val || '' })}
              onClear={() => updateFilter(field.internalName, { ...current, text: '' })}
            />
          </div>
        );
      case 'Choice': {
        const options: IDropdownOption[] = [
          { key: '', text: `All ${field.displayName}` },
          ...(field.choices || []).map(c => ({ key: c, text: c }))
        ];
        return (
          <div key={field.internalName} className={styles.filterField}>
            <Dropdown
              placeholder={`All ${field.displayName}`}
              options={options}
              selectedKey={current.choice || ''}
              onChange={(_, opt) => updateFilter(field.internalName, { ...current, choice: (opt?.key as string) || '' })}
            />
          </div>
        );
      }
      case 'DateTime':
        return (
          <div key={field.internalName} className={styles.dateRangeField}>
            <div style={{ flex: 1 }}>
              <Label>{`${field.displayName} From`}</Label>
              <DatePicker
                placeholder="From date"
                value={current.dateFrom ? new Date(current.dateFrom) : undefined}
                onSelectDate={date => updateFilter(field.internalName, {
                  ...current,
                  dateFrom: date ? date.toISOString().split('T')[0] : undefined
                })}
              />
            </div>
            <div style={{ flex: 1 }}>
              <Label>{`${field.displayName} To`}</Label>
              <DatePicker
                placeholder="To date"
                value={current.dateTo ? new Date(current.dateTo) : undefined}
                onSelectDate={date => updateFilter(field.internalName, {
                  ...current,
                  dateTo: date ? date.toISOString().split('T')[0] : undefined
                })}
              />
            </div>
          </div>
        );
      case 'Number':
        return (
          <div key={field.internalName} className={styles.numberRangeField}>
            <TextField
              label={`${field.displayName} Min`}
              type="number"
              value={current.numberMin !== undefined ? String(current.numberMin) : ''}
              onChange={(_, val) => updateFilter(field.internalName, {
                ...current,
                numberMin: val !== '' && val !== undefined ? Number(val) : undefined
              })}
              styles={{ root: { flex: 1 } }}
            />
            <TextField
              label={`${field.displayName} Max`}
              type="number"
              value={current.numberMax !== undefined ? String(current.numberMax) : ''}
              onChange={(_, val) => updateFilter(field.internalName, {
                ...current,
                numberMax: val !== '' && val !== undefined ? Number(val) : undefined
              })}
              styles={{ root: { flex: 1 } }}
            />
          </div>
        );
      case 'Boolean':
        return (
          <div key={field.internalName} className={styles.filterField}>
            <Label>{field.displayName}</Label>
            <Dropdown
              options={[
                { key: '', text: 'Any' },
                { key: 'true', text: 'Yes' },
                { key: 'false', text: 'No' }
              ]}
              selectedKey={current.boolValue === true ? 'true' : current.boolValue === false ? 'false' : ''}
              onChange={(_, opt) => {
                const boolValue = opt?.key === 'true' ? true : opt?.key === 'false' ? false : undefined;
                updateFilter(field.internalName, { ...current, boolValue });
              }}
            />
          </div>
        );
      default:
        return <React.Fragment key={field.internalName} />;
    }
  };

  return (
    <div className={styles.filterBar}>
      {filterableFields.map(f => renderFilterControl(f))}
      {hasFilters && (
        <div className={styles.clearButton}>
          <DefaultButton text="Clear Filters" iconProps={{ iconName: 'ClearFilter' }} onClick={clearAll} />
        </div>
      )}
    </div>
  );
};

export default FilterBar;
