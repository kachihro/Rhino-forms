import * as React from 'react';
import { useState } from 'react';
import {
  TextField, Dropdown, IDropdownOption, Toggle, IconButton, PrimaryButton, DefaultButton, Text, MessageBar, MessageBarType
} from '@fluentui/react';
import { IRhinoConfig, IFieldDefinition, FieldType, FIELD_TYPES, NumberFormat } from '../../../../shared/types/RhinoConfig';
import { toInternalName, isValidInternalName, buildDefaultColumn } from '../../../../shared/config/helpers';
import { IValidationIssue } from '../../../../shared/config/validate';
import { withConfig, moveItem } from '../DesignerApp';
import styles from '../DesignerApp.module.scss';

export interface IFieldsStepProps {
  config: IRhinoConfig;
  onChange: (next: IRhinoConfig) => void;
  issues: IValidationIssue[];
}

const TYPE_LABELS: Record<FieldType, string> = {
  Text: 'Single line of text',
  Note: 'Multiple lines of text',
  Choice: 'Choice',
  MultiChoice: 'Choice (multiple)',
  DateTime: 'Date',
  Number: 'Number',
  Boolean: 'Yes / No',
  User: 'Person',
  URL: 'Hyperlink'
};

const NUMBER_FORMATS: IDropdownOption[] = [
  { key: 'decimal', text: 'Decimal' },
  { key: 'integer', text: 'Whole number' },
  { key: 'currency', text: 'Currency (AUD)' },
  { key: 'percent', text: 'Percent' }
];

export const FieldsStep: React.FC<IFieldsStepProps> = ({ config, onChange, issues }) => {
  const [open, setOpen] = useState<string | undefined>();
  const [newName, setNewName] = useState('');
  const [newType, setNewType] = useState<FieldType>('Text');
  const [lockedNames, setLockedNames] = useState<Record<string, boolean>>({});

  const typeOptions: IDropdownOption[] = FIELD_TYPES.map(t => ({ key: t, text: TYPE_LABELS[t] }));
  const issueFor = (index: number, prop: string): string | undefined =>
    issues.filter(i => i.path === `fields[${index}].${prop}` && i.level === 'error')[0]?.message;

  const updateField = (index: number, mutate: (f: IFieldDefinition) => void): void => {
    onChange(withConfig(config, d => mutate(d.fields[index])));
  };

  const renameInternal = (index: number, oldName: string, newInternal: string): void => {
    onChange(withConfig(config, d => {
      d.fields[index].internalName = newInternal;
      d.form.sections.forEach(s => s.fields.forEach(p => { if (p.field === oldName) p.field = newInternal; }));
      d.grid.columns.forEach(c => { if (c.field === oldName) c.field = newInternal; });
      if (d.grid.defaultSort?.field === oldName) d.grid.defaultSort.field = newInternal;
    }));
  };

  const setDisplayName = (index: number, value: string): void => {
    const f = config.fields[index];
    const locked = lockedNames[f.internalName] || f.internalName === 'Title';
    if (locked) { updateField(index, x => { x.displayName = value; }); return; }
    const nextInternal = toInternalName(value);
    onChange(withConfig(config, d => {
      const old = d.fields[index].internalName;
      d.fields[index].displayName = value;
      if (nextInternal && nextInternal !== old && !d.fields.some((x, i) => i !== index && x.internalName.toLowerCase() === nextInternal.toLowerCase())) {
        d.fields[index].internalName = nextInternal;
        d.form.sections.forEach(s => s.fields.forEach(p => { if (p.field === old) p.field = nextInternal; }));
        d.grid.columns.forEach(c => { if (c.field === old) c.field = nextInternal; });
        if (d.grid.defaultSort?.field === old) d.grid.defaultSort.field = nextInternal;
      }
    }));
  };

  const addField = (): void => {
    const displayName = newName.trim();
    if (!displayName) return;
    let internal = toInternalName(displayName);
    let n = 2;
    while (config.fields.some(f => f.internalName.toLowerCase() === internal.toLowerCase())) internal = `${toInternalName(displayName)}${n++}`;
    const field: IFieldDefinition = { internalName: internal, displayName, type: newType };
    if (newType === 'Choice' || newType === 'MultiChoice') field.choices = ['Option 1', 'Option 2'];
    if (newType === 'DateTime') field.dateOnly = true;
    onChange(withConfig(config, d => {
      d.fields.push(field);
      const last = d.form.sections[d.form.sections.length - 1];
      if (last) last.fields.push({ field: internal, width: newType === 'Note' ? 'full' : 'half' });
      d.grid.columns.push(buildDefaultColumn(field));
    }));
    setNewName('');
    setOpen(internal);
  };

  const removeField = (index: number): void => {
    const f = config.fields[index];
    if (f.internalName === 'Title') return;
    if (!window.confirm(`Remove "${f.displayName}" from the design? (The SharePoint column, if provisioned, is not deleted.)`)) return;
    onChange(withConfig(config, d => {
      d.fields.splice(index, 1);
      d.form.sections.forEach(s => { s.fields = s.fields.filter(p => p.field !== f.internalName); });
      d.grid.columns = d.grid.columns.filter(c => c.field !== f.internalName);
      if (d.grid.defaultSort?.field === f.internalName) d.grid.defaultSort = { field: 'ID', direction: 'desc' };
    }));
  };

  const move = (index: number, delta: number): void => {
    onChange(withConfig(config, d => { d.fields = moveItem(d.fields, index, index + delta); }));
  };

  const renderDetails = (f: IFieldDefinition, index: number): JSX.Element => (
    <div className={styles.fieldDetails}>
      <div className={styles.formRow}>
        <TextField label="Description / help text" value={f.description || ''} onChange={(_, v) => updateField(index, x => { x.description = v || undefined; })} />
        {(f.type === 'Text' || f.type === 'Choice' || f.type === 'Number') && (
          <TextField label="Default value" value={f.defaultValue === undefined ? '' : String(f.defaultValue)}
            onChange={(_, v) => updateField(index, x => { x.defaultValue = v === '' ? undefined : (f.type === 'Number' ? Number(v) : v); })} />
        )}
        {f.type === 'Boolean' && (
          <Toggle label="Default" checked={!!f.defaultValue} onText="Yes" offText="No" onChange={(_, c) => updateField(index, x => { x.defaultValue = !!c; })} />
        )}
        {f.type === 'DateTime' && (
          <>
            <Dropdown label="Format" options={[{ key: 'date', text: 'Date only' }, { key: 'datetime', text: 'Date and time' }]} selectedKey={f.dateOnly === false ? 'datetime' : 'date'}
              onChange={(_, o) => updateField(index, x => { x.dateOnly = o?.key !== 'datetime'; })} />
            <Toggle label="Default to today" checked={f.defaultValue === 'today'} onChange={(_, c) => updateField(index, x => { x.defaultValue = c ? 'today' : undefined; })} />
          </>
        )}
        {f.type === 'Text' && (
          <TextField label="Max length" type="number" value={String(f.maxLength || 255)} onChange={(_, v) => updateField(index, x => { x.maxLength = v ? Number(v) : undefined; })} />
        )}
        {f.type === 'Number' && (
          <>
            <Dropdown label="Number format" options={NUMBER_FORMATS} selectedKey={f.numberFormat || 'decimal'} onChange={(_, o) => updateField(index, x => { x.numberFormat = o?.key as NumberFormat; })} />
            <TextField label="Decimals" type="number" value={f.decimals === undefined ? '' : String(f.decimals)} placeholder="auto" onChange={(_, v) => updateField(index, x => { x.decimals = v === '' ? undefined : Number(v); })} />
            <TextField label="Min" type="number" value={f.min === undefined ? '' : String(f.min)} onChange={(_, v) => updateField(index, x => { x.min = v === '' ? undefined : Number(v); })} />
            <TextField label="Max" type="number" value={f.max === undefined ? '' : String(f.max)} onChange={(_, v) => updateField(index, x => { x.max = v === '' ? undefined : Number(v); })} />
          </>
        )}
        {f.type === 'User' && (
          <Toggle label="Allow multiple people" checked={!!f.allowMultiple} onChange={(_, c) => updateField(index, x => { x.allowMultiple = !!c; })} />
        )}
      </div>
      {(f.type === 'Choice' || f.type === 'MultiChoice') && (
        <TextField label="Choices (one per line)" multiline rows={Math.min(8, Math.max(3, (f.choices || []).length + 1))}
          value={(f.choices || []).join('\n')}
          errorMessage={issueFor(index, 'choices')}
          onChange={(_, v) => updateField(index, x => { x.choices = (v || '').split('\n').map(s => s.trim()).filter(Boolean); })}
          description="Tip: the grid colours status-like values automatically (e.g. Complete = green, On Hold = grey)." />
      )}
    </div>
  );

  return (
    <div>
      <div className={styles.stepHeader}>
        <h2>Define the fields</h2>
        <p>Display names are what people see. Internal names are what SharePoint stores – they are generated for you and never contain spaces or <span className={styles.mono}>_x0020_</span>.</p>
      </div>

      {config.fields.some(f => !isValidInternalName(f.internalName)) && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.notice}>One or more internal names are invalid. Letters, digits and underscores only, starting with a letter.</MessageBar>
      )}

      <div className={styles.fieldRow} style={{ background: '#fafafa', border: 'none' }}>
        <span />
        <span className={styles.colHead}>Display name</span>
        <span className={styles.colHead}>Internal name</span>
        <span className={styles.colHead}>Type</span>
        <span className={styles.colHead}>Required</span>
        <span />
      </div>

      {config.fields.map((f, index) => {
        const isTitle = f.internalName === 'Title';
        const isOpen = open === f.internalName;
        return (
          <div key={`${index}-${f.internalName}`}>
            <div className={`${styles.fieldRow} ${isOpen ? styles.fieldRowOpen : ''}`}>
              <span className={styles.fieldIndex}>{index + 1}</span>
              <TextField value={f.displayName} onChange={(_, v) => setDisplayName(index, v || '')} errorMessage={issueFor(index, 'displayName')} />
              <TextField value={f.internalName} disabled={isTitle} className={styles.mono}
                onChange={(_, v) => { setLockedNames(prev => ({ ...prev, [v || '']: true })); renameInternal(index, f.internalName, (v || '').replace(/[^A-Za-z0-9_]/g, '')); }}
                errorMessage={issueFor(index, 'internalName')} />
              <Dropdown options={typeOptions} selectedKey={f.type} disabled={isTitle}
                onChange={(_, o) => updateField(index, x => {
                  x.type = o?.key as FieldType;
                  if ((x.type === 'Choice' || x.type === 'MultiChoice') && !x.choices?.length) x.choices = ['Option 1', 'Option 2'];
                  if (x.type === 'DateTime' && x.dateOnly === undefined) x.dateOnly = true;
                })} />
              <Toggle checked={!!f.required} disabled={isTitle} onChange={(_, c) => updateField(index, x => { x.required = !!c; })} styles={{ root: { marginBottom: 0 } }} />
              <div className={styles.rowBtns}>
                <IconButton iconProps={{ iconName: isOpen ? 'ChevronUp' : 'Settings' }} title="Details" className={styles.iconBtn} onClick={() => setOpen(isOpen ? undefined : f.internalName)} />
                <IconButton iconProps={{ iconName: 'Up' }} title="Move up" className={styles.iconBtn} disabled={index === 0} onClick={() => move(index, -1)} />
                <IconButton iconProps={{ iconName: 'Down' }} title="Move down" className={styles.iconBtn} disabled={index === config.fields.length - 1} onClick={() => move(index, 1)} />
                <IconButton iconProps={{ iconName: 'Delete' }} title={isTitle ? 'Title cannot be removed' : 'Remove'} className={styles.iconBtn} disabled={isTitle} onClick={() => removeField(index)} />
              </div>
            </div>
            {isOpen && renderDetails(f, index)}
          </div>
        );
      })}

      <div className={styles.addRow}>
        <TextField placeholder="New field display name" value={newName} onChange={(_, v) => setNewName(v || '')} styles={{ root: { width: 260 } }}
          onKeyDown={e => { if (e.key === 'Enter') addField(); }} />
        <Dropdown options={typeOptions} selectedKey={newType} onChange={(_, o) => setNewType(o?.key as FieldType)} styles={{ root: { width: 200 } }} />
        <PrimaryButton text="Add field" iconProps={{ iconName: 'Add' }} onClick={addField} disabled={!newName.trim()} />
        {newName.trim() && <Text variant="small" className={styles.muted}>→ <span className={styles.mono}>{toInternalName(newName)}</span></Text>}
      </div>

      <div className={styles.docsBox} style={{ marginTop: 20 }}>
        <strong>Good to know</strong> · Changing an internal name after the list is provisioned creates a <em>new</em> column; the old one stays in SharePoint.
        Provisioning never deletes columns. Required, description and choices are kept in sync each time you provision.
        <DefaultButton text="Restore Title field" style={{ marginLeft: 12 }} iconProps={{ iconName: 'Undo' }}
          disabled={config.fields.some(f => f.internalName === 'Title')}
          onClick={() => onChange(withConfig(config, d => {
            d.fields.unshift({ internalName: 'Title', displayName: 'Title', type: 'Text', required: true });
            d.form.sections[0]?.fields.unshift({ field: 'Title', width: 'full' });
            d.grid.columns.unshift({ field: 'Title', sortable: true, filterable: true, width: 200 });
          }))} />
      </div>
    </div>
  );
};

export default FieldsStep;
