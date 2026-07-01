import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Toggle,
  DatePicker,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { IFieldDefinition, IGridItem } from '../../types/IFieldConfig';
import { GridDataService } from '../../services/GridDataService';
import styles from './ItemForm.module.scss';

export interface IItemFormProps {
  mode: 'add' | 'edit';
  fields: IFieldDefinition[];
  item?: IGridItem;
  listName: string;
  dataService: GridDataService;
  onSaved: () => void;
  onDismiss: () => void;
}

type FormData = Record<string, unknown>;

export const ItemForm: React.FC<IItemFormProps> = ({ mode, fields, item, listName, dataService, onSaved, onDismiss }) => {
  const [formData, setFormData] = useState<FormData>({});
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [validationErrors, setValidationErrors] = useState<Record<string, string>>({});

  useEffect(() => {
    if (mode === 'edit' && item) {
      const data: FormData = {};
      for (const field of fields) {
        if (field.type === 'User') {
          const userValue = item[field.internalName] as { Id?: number; Title?: string } | undefined;
          data[`${field.internalName}Id`] = userValue?.Id;
          data[`_display_${field.internalName}`] = userValue?.Title || '';
        } else if (field.type === 'DateTime') {
          const raw = item[field.internalName];
          data[field.internalName] = raw ? new Date(raw as string) : null;
        } else {
          data[field.internalName] = item[field.internalName];
        }
      }
      setFormData(data);
    } else {
      const defaults: FormData = {};
      for (const field of fields) {
        if (field.defaultValue !== undefined) {
          defaults[field.internalName] = field.defaultValue;
        }
      }
      setFormData(defaults);
    }
  }, [mode, item, fields]);

  const validate = (): boolean => {
    const errors: Record<string, string> = {};
    for (const field of fields) {
      if (field.required) {
        const val = field.type === 'User' ? formData[`${field.internalName}Id`] : formData[field.internalName];
        if (val === undefined || val === null || val === '') {
          errors[field.internalName] = `${field.displayName} is required`;
        }
      }
    }
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setSaving(true);
    setError(undefined);
    try {
      const payload: Record<string, unknown> = {};
      for (const field of fields) {
        if (field.type === 'User') {
          const userId = formData[`${field.internalName}Id`];
          if (userId !== undefined) {
            payload[`${field.internalName}Id`] = userId;
          }
        } else if (field.type === 'DateTime') {
          const dateVal = formData[field.internalName] as Date | null;
          payload[field.internalName] = dateVal ? dateVal.toISOString() : null;
        } else if (field.type === 'URL') {
          payload[field.internalName] = {
            Description: formData[`${field.internalName}_desc`] || '',
            Url: formData[field.internalName] || ''
          };
        } else {
          payload[field.internalName] = formData[field.internalName];
        }
      }

      if (mode === 'add') {
        await dataService.addItem(listName, payload);
      } else if (item) {
        await dataService.updateItem(listName, item.Id, payload);
      }
      onSaved();
    } catch (e) {
      setError(`Failed to save: ${e.message}`);
    } finally {
      setSaving(false);
    }
  };

  const setField = (internalName: string, value: unknown): void => {
    setFormData(prev => ({ ...prev, [internalName]: value }));
    if (validationErrors[internalName]) {
      setValidationErrors(prev => {
        const next = { ...prev };
        delete next[internalName];
        return next;
      });
    }
  };

  const renderField = (field: IFieldDefinition): JSX.Element => {
    const errorMessage = validationErrors[field.internalName];

    switch (field.type) {
      case 'Text':
        return (
          <TextField
            key={field.internalName}
            label={field.displayName}
            required={field.required}
            value={(formData[field.internalName] as string) || ''}
            onChange={(_, val) => setField(field.internalName, val)}
            errorMessage={errorMessage}
            className={styles.formField}
            maxLength={field.maxLength || 255}
          />
        );
      case 'Note':
        return (
          <TextField
            key={field.internalName}
            label={field.displayName}
            required={field.required}
            multiline
            rows={4}
            value={(formData[field.internalName] as string) || ''}
            onChange={(_, val) => setField(field.internalName, val)}
            errorMessage={errorMessage}
            className={styles.formField}
          />
        );
      case 'Choice': {
        const options: IDropdownOption[] = (field.choices || []).map(c => ({ key: c, text: c }));
        return (
          <Dropdown
            key={field.internalName}
            label={field.displayName}
            required={field.required}
            options={options}
            selectedKey={(formData[field.internalName] as string) || null}
            onChange={(_, opt) => setField(field.internalName, opt?.key)}
            errorMessage={errorMessage}
            className={styles.formField}
          />
        );
      }
      case 'DateTime':
        return (
          <DatePicker
            key={field.internalName}
            label={field.displayName}
            isRequired={field.required}
            value={(formData[field.internalName] as Date) || undefined}
            onSelectDate={date => setField(field.internalName, date)}
            className={styles.formField}
          />
        );
      case 'Number':
        return (
          <TextField
            key={field.internalName}
            label={field.displayName}
            required={field.required}
            type="number"
            value={formData[field.internalName] !== undefined ? String(formData[field.internalName]) : ''}
            onChange={(_, val) => setField(field.internalName, val !== '' ? Number(val) : undefined)}
            errorMessage={errorMessage}
            className={styles.formField}
          />
        );
      case 'Boolean':
        return (
          <Toggle
            key={field.internalName}
            label={field.displayName}
            checked={!!formData[field.internalName]}
            onChange={(_, checked) => setField(field.internalName, checked)}
            className={styles.formField}
          />
        );
      case 'User':
        return (
          <TextField
            key={field.internalName}
            label={`${field.displayName} (User ID)`}
            required={field.required}
            type="number"
            placeholder="Enter SharePoint user ID"
            value={formData[`${field.internalName}Id`] !== undefined ? String(formData[`${field.internalName}Id`]) : ''}
            onChange={(_, val) => setField(`${field.internalName}Id`, val !== '' ? Number(val) : undefined)}
            errorMessage={errorMessage}
            className={styles.formField}
          />
        );
      case 'URL':
        return (
          <div key={field.internalName} className={styles.formField}>
            <TextField
              label={`${field.displayName} (URL)`}
              required={field.required}
              value={(formData[field.internalName] as string) || ''}
              onChange={(_, val) => setField(field.internalName, val)}
              errorMessage={errorMessage}
              placeholder="https://..."
            />
            <TextField
              label={`${field.displayName} (Description)`}
              value={(formData[`${field.internalName}_desc`] as string) || ''}
              onChange={(_, val) => setField(`${field.internalName}_desc`, val)}
              placeholder="Link description"
            />
          </div>
        );
      default:
        return (
          <TextField
            key={field.internalName}
            label={field.displayName}
            required={field.required}
            value={(formData[field.internalName] as string) || ''}
            onChange={(_, val) => setField(field.internalName, val)}
            className={styles.formField}
          />
        );
    }
  };

  const listDisplayName = listName.replace(/_/g, ' ');
  const panelTitle = mode === 'add' ? `Add ${listDisplayName}` : `Edit ${listDisplayName}`;

  return (
    <Panel
      isOpen
      type={PanelType.medium}
      headerText={panelTitle}
      onDismiss={onDismiss}
      onRenderFooterContent={() => (
        <div className={styles.formActions}>
          <PrimaryButton text="Save" onClick={handleSave} disabled={saving} />
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        </div>
      )}
      isFooterAtBottom
    >
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      {saving && <Spinner size={SpinnerSize.small} label="Saving..." />}
      {fields.map(field => renderField(field))}
    </Panel>
  );
};

export default ItemForm;
