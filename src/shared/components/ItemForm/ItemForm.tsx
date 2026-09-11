import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Panel, PanelType, Dialog, DialogType, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Spinner, SpinnerSize, Icon, Text
} from '@fluentui/react';
import { IRhinoConfig, IGridItem, IFieldDefinition, IFormSection, IFormFieldPlacement } from '../../types/RhinoConfig';
import { fieldByName, itemLabels } from '../../config/helpers';
import { PeopleService } from '../../services/PeopleService';
import FieldEditor from '../fields/FieldEditor';
import { FormValues, defaultValues, itemToValues, valuesToPayload, validateValues } from '../fields/formValues';
import { formatDate, userArray } from '../cells/format';
import styles from './ItemForm.module.scss';

export type FormMode = 'add' | 'edit' | 'view';

export interface IItemFormProps {
  config: IRhinoConfig;
  mode: FormMode;
  item?: IGridItem;
  peopleService?: PeopleService;
  onSave: (payload: Record<string, unknown>) => Promise<void>;
  onDismiss: () => void;
  onEdit?: () => void;
  onDelete?: () => void;
  /** Render inline (no panel/dialog chrome). Used by the designer preview. */
  inline?: boolean;
}

function widthClass(width: IFormFieldPlacement['width'], columns: number): string {
  if (columns === 1) return styles.wFull;
  switch (width) {
    case 'full': return styles.wFull;
    case 'third': return columns === 3 ? styles.wThird : styles.wHalf;
    case 'twoThirds': return columns === 3 ? styles.wTwoThirds : styles.wFull;
    case 'half':
    default: return columns === 3 ? styles.wThird : styles.wHalf;
  }
}

function titleFor(template: string | undefined, fallback: string, config: IRhinoConfig, item?: IGridItem): string {
  const { singular } = itemLabels(config);
  const t = template || fallback;
  return t
    .replace(/\{item\}/gi, singular)
    .replace(/\{Title\}/g, item?.Title ? String(item.Title) : singular)
    .replace(/\{Id\}/gi, item ? String(item.Id) : '');
}

export const ItemForm: React.FC<IItemFormProps> = ({ config, mode, item, peopleService, onSave, onDismiss, onEdit, onDelete, inline }) => {
  const editableFields = useMemo(() => config.fields.filter(f => !f.isSystem), [config.fields]);
  const [values, setValues] = useState<FormValues>({});
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | undefined>();
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  const [dirty, setDirty] = useState(false);

  useEffect(() => {
    setValues(mode === 'add' || !item ? defaultValues(editableFields) : itemToValues(item, editableFields));
    setErrors({});
    setSaveError(undefined);
    setDirty(false);
    const c: Record<string, boolean> = {};
    config.form.sections.forEach(s => { if (s.collapsible && s.collapsedByDefault) c[s.id] = true; });
    setCollapsed(c);
  }, [mode, item, editableFields, config.form.sections]);

  const isEditable = (name: string): boolean => {
    if (mode === 'view') return false;
    const f = fieldByName(config, name);
    if (!f || f.isSystem) return false;
    let readOnly = false;
    config.form.sections.forEach(s => s.fields.forEach(p => { if (p.field === name && p.readOnly) readOnly = true; }));
    return !readOnly;
  };

  const setField = (name: string, v: unknown): void => {
    setValues(prev => ({ ...prev, [name]: v }));
    setDirty(true);
    if (errors[name]) {
      setErrors(prev => { const n = { ...prev }; delete n[name]; return n; });
    }
  };

  const handleSave = async (): Promise<void> => {
    const errs = validateValues(values, editableFields, isEditable);
    setErrors(errs);
    if (Object.keys(errs).length > 0) {
      // open any collapsed section containing an error
      const open: Record<string, boolean> = { ...collapsed };
      config.form.sections.forEach(s => { if (s.fields.some(p => errs[p.field])) open[s.id] = false; });
      setCollapsed(open);
      return;
    }
    setSaving(true);
    setSaveError(undefined);
    try {
      await onSave(valuesToPayload(values, editableFields, isEditable));
    } catch (e) {
      setSaveError((e as Error).message || 'Save failed');
    } finally {
      setSaving(false);
    }
  };

  const renderPlacement = (section: IFormSection, p: IFormFieldPlacement): React.ReactNode => {
    const field: IFieldDefinition | undefined = fieldByName(config, p.field);
    if (!field) return undefined;
    const value = field.isSystem && item ? item[field.internalName] : values[field.internalName];
    return (
      <div key={`${p.field}-${item ? item.Id : 'new'}`} className={`${styles.cell} ${widthClass(p.width, section.columns || 2)}`}>
        <FieldEditor
          field={field}
          placement={p}
          value={value}
          onChange={v => setField(field.internalName, v)}
          errorMessage={errors[field.internalName]}
          readOnly={!isEditable(field.internalName)}
          peopleService={peopleService}
        />
      </div>
    );
  };

  const renderSection = (section: IFormSection): React.ReactNode => {
    if (section.fields.length === 0) return undefined;
    const isCollapsed = !!collapsed[section.id];
    const showHeader = !!section.title || !!section.description || !!section.collapsible;
    return (
      <section key={section.id} className={styles.section}>
        {showHeader && (
          <div
            className={`${styles.sectionHeader} ${section.collapsible ? styles.clickable : ''}`}
            onClick={section.collapsible ? () => setCollapsed(prev => ({ ...prev, [section.id]: !isCollapsed })) : undefined}
          >
            <div>
              {section.title && <Text className={styles.sectionTitle}>{section.title}</Text>}
              {section.description && <Text variant="small" className={styles.sectionDesc}>{section.description}</Text>}
            </div>
            {section.collapsible && <Icon iconName={isCollapsed ? 'ChevronDown' : 'ChevronUp'} className={styles.chev} />}
          </div>
        )}
        {!isCollapsed && (
          <div className={`${styles.grid} ${section.columns === 3 ? styles.cols3 : section.columns === 1 ? styles.cols1 : styles.cols2}`}>
            {section.fields.map(p => renderPlacement(section, p))}
          </div>
        )}
      </section>
    );
  };

  const errorCount = Object.keys(errors).length;
  const { singular } = itemLabels(config);
  const title = mode === 'add'
    ? titleFor(config.form.addTitle, 'New {item}', config)
    : mode === 'edit'
      ? titleFor(config.form.editTitle, 'Edit {Title}', config, item)
      : (item?.Title ? String(item.Title) : `${singular} #${item?.Id || ''}`);

  const body = (
    <div className={styles.body}>
      {saveError && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setSaveError(undefined)}>{saveError}</MessageBar>
      )}
      {errorCount > 0 && (
        <MessageBar messageBarType={MessageBarType.warning}>
          {errorCount === 1 ? 'One field needs attention.' : `${errorCount} fields need attention.`}
        </MessageBar>
      )}
      {config.form.sections.map(renderSection)}
      {config.form.showSystemInfo && item && mode !== 'add' && (
        <div className={styles.systemInfo}>
          <span><Icon iconName="Add" /> Created {formatDate(item.Created, true)}{userArray(item.Author)[0]?.Title ? ` by ${userArray(item.Author)[0].Title}` : ''}</span>
          <span><Icon iconName="Edit" /> Modified {formatDate(item.Modified, true)}{userArray(item.Editor)[0]?.Title ? ` by ${userArray(item.Editor)[0].Title}` : ''}</span>
          <span><Icon iconName="NumberSymbol" /> ID {item.Id}</span>
        </div>
      )}
    </div>
  );

  const footer = (
    <div className={styles.footer}>
      <div className={styles.footerLeft}>
        {mode === 'view' && onEdit && config.grid.allowEdit !== false && (
          <PrimaryButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={onEdit} />
        )}
        {mode !== 'view' && (
          <PrimaryButton text={saving ? 'Saving…' : (config.form.saveLabel || 'Save')} iconProps={{ iconName: 'Save' }} onClick={() => { void handleSave(); }} disabled={saving} />
        )}
        <DefaultButton text={mode === 'view' ? 'Close' : (config.form.cancelLabel || 'Cancel')} onClick={onDismiss} disabled={saving} />
        {saving && <Spinner size={SpinnerSize.small} />}
      </div>
      {mode !== 'add' && onDelete && config.grid.allowDelete !== false && (
        <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} className={styles.deleteBtn} onClick={onDelete} disabled={saving} />
      )}
    </div>
  );

  if (inline) {
    return (
      <div className={styles.inlineFrame}>
        <div className={styles.inlineHeader}><Text className={styles.inlineTitle}>{title}</Text></div>
        {body}
        {footer}
      </div>
    );
  }

  const confirmDismiss = (): void => {
    if (dirty && mode !== 'view' && !window.confirm('Discard unsaved changes?')) return;
    onDismiss();
  };

  if (config.form.surface === 'dialog') {
    const width = config.form.size === 'extraLarge' ? 1100 : config.form.size === 'large' ? 860 : 640;
    return (
      <Dialog
        hidden={false}
        onDismiss={confirmDismiss}
        dialogContentProps={{ type: DialogType.normal, title, showCloseButton: true }}
        modalProps={{ isBlocking: true, styles: { main: { maxWidth: width, width: '92vw', selectors: { '@media (min-width: 480px)': { maxWidth: width, minWidth: Math.min(width, 480) } } } } }}
      >
        {body}
        {footer}
      </Dialog>
    );
  }

  const panelType = config.form.size === 'extraLarge' ? PanelType.extraLarge : config.form.size === 'large' ? PanelType.large : PanelType.medium;
  return (
    <Panel
      isOpen
      type={panelType}
      headerText={title}
      onDismiss={confirmDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => footer}
      closeButtonAriaLabel="Close"
      styles={{ content: { paddingBottom: 24 }, footerInner: { padding: '12px 24px' } }}
    >
      {body}
    </Panel>
  );
};

export default ItemForm;
