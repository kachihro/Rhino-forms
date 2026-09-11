import * as React from 'react';
import { useMemo, useState } from 'react';
import {
  TextField, Dropdown, IDropdownOption, Toggle, IconButton, DefaultButton, Text, ChoiceGroup
} from '@fluentui/react';
import { IRhinoConfig, IFormSection, IFormFieldPlacement, FieldWidth, FormSurface, FormSize } from '../../../../shared/types/RhinoConfig';
import { fieldByName, newSectionId, systemFields } from '../../../../shared/config/helpers';
import { sampleItems } from '../../../../shared/config/sample';
import ItemForm from '../../../../shared/components/ItemForm/ItemForm';
import { withConfig, moveItem } from '../DesignerApp';
import styles from '../DesignerApp.module.scss';

export interface IFormStepProps {
  config: IRhinoConfig;
  onChange: (next: IRhinoConfig) => void;
}

const WIDTHS: IDropdownOption[] = [
  { key: 'full', text: 'Full width' },
  { key: 'half', text: 'Half' },
  { key: 'third', text: 'Third' },
  { key: 'twoThirds', text: 'Two thirds' }
];

export const FormStep: React.FC<IFormStepProps> = ({ config, onChange }) => {
  const [previewMode, setPreviewMode] = useState<'add' | 'edit' | 'view'>('edit');
  const sample = useMemo(() => sampleItems(config, 1)[0], [config]);

  const placedNames: Record<string, boolean> = {};
  config.form.sections.forEach(s => s.fields.forEach(p => { placedNames[p.field] = true; }));
  const unplaced = config.fields.filter(f => !f.isSystem && !placedNames[f.internalName]);
  const unplacedSystem = systemFields().filter(f => !placedNames[f.internalName]);

  const updateSection = (si: number, mutate: (s: IFormSection) => void): void =>
    onChange(withConfig(config, d => mutate(d.form.sections[si])));

  const updatePlacement = (si: number, pi: number, mutate: (p: IFormFieldPlacement) => void): void =>
    onChange(withConfig(config, d => mutate(d.form.sections[si].fields[pi])));

  const addSection = (): void =>
    onChange(withConfig(config, d => { d.form.sections.push({ id: newSectionId(), title: `Section ${d.form.sections.length + 1}`, columns: 2, fields: [] }); }));

  const removeSection = (si: number): void =>
    onChange(withConfig(config, d => {
      const [removed] = d.form.sections.splice(si, 1);
      const target = d.form.sections[Math.max(0, si - 1)];
      if (target) target.fields = target.fields.concat(removed.fields);
      else d.form.sections.push({ id: newSectionId(), title: '', columns: 2, fields: removed.fields });
    }));

  const moveSection = (si: number, delta: number): void =>
    onChange(withConfig(config, d => { d.form.sections = moveItem(d.form.sections, si, si + delta); }));

  const movePlacement = (si: number, pi: number, delta: number): void =>
    updateSection(si, s => { s.fields = moveItem(s.fields, pi, pi + delta); });

  const moveToSection = (si: number, pi: number, targetId: string): void =>
    onChange(withConfig(config, d => {
      const [p] = d.form.sections[si].fields.splice(pi, 1);
      const target = d.form.sections.filter(s => s.id === targetId)[0];
      if (target) target.fields.push(p);
    }));

  const removePlacement = (si: number, pi: number): void =>
    updateSection(si, s => { s.fields.splice(pi, 1); });

  const place = (name: string): void =>
    onChange(withConfig(config, d => {
      const f = fieldByName(d, name);
      const last = d.form.sections[d.form.sections.length - 1];
      if (last) last.fields.push({ field: name, width: f?.type === 'Note' ? 'full' : 'half', readOnly: !!f?.isSystem });
    }));

  const sectionOptions: IDropdownOption[] = config.form.sections.map((s, i) => ({ key: s.id, text: s.title || `Section ${i + 1}` }));

  return (
    <div>
      <div className={styles.stepHeader}>
        <h2>Lay out the form</h2>
        <p>Group fields into sections, choose column counts and widths, and override labels or placeholders. The preview on the right updates live.</p>
      </div>

      <div className={styles.split}>
        <div>
          <div className={styles.panel} style={{ marginBottom: 12 }}>
            <div className={styles.panelTitle}>Form settings</div>
            <div className={styles.settingsGrid}>
              <ChoiceGroup label="Opens as" selectedKey={config.form.surface}
                options={[{ key: 'panel', text: 'Side panel' }, { key: 'dialog', text: 'Dialog' }]}
                onChange={(_, o) => onChange(withConfig(config, d => { d.form.surface = o?.key as FormSurface; }))} />
              <Dropdown label="Size" selectedKey={config.form.size} options={[{ key: 'medium', text: 'Medium' }, { key: 'large', text: 'Large' }, { key: 'extraLarge', text: 'Extra large' }]}
                onChange={(_, o) => onChange(withConfig(config, d => { d.form.size = o?.key as FormSize; }))} />
              <TextField label="New item title" value={config.form.addTitle || ''} placeholder="New {item}"
                description="{item} = item label" onChange={(_, v) => onChange(withConfig(config, d => { d.form.addTitle = v || ''; }))} />
              <TextField label="Edit title" value={config.form.editTitle || ''} placeholder="Edit {Title}"
                description="{Title} = the item's title, {Id} = ID" onChange={(_, v) => onChange(withConfig(config, d => { d.form.editTitle = v || ''; }))} />
            </div>
          </div>

          {config.form.sections.map((section, si) => (
            <div key={section.id} className={styles.section}>
              <div className={styles.sectionHead}>
                <TextField label="Section title" value={section.title || ''} placeholder="(no heading)" onChange={(_, v) => updateSection(si, s => { s.title = v || ''; })} />
                <Dropdown label="Columns" selectedKey={section.columns || 2} options={[{ key: 1, text: '1' }, { key: 2, text: '2' }, { key: 3, text: '3' }]}
                  onChange={(_, o) => updateSection(si, s => { s.columns = Number(o?.key) as 1 | 2 | 3; })} />
                <Toggle label="Collapsible" checked={!!section.collapsible} onChange={(_, c) => updateSection(si, s => { s.collapsible = !!c; if (!c) s.collapsedByDefault = false; })} />
                <div className={styles.rowBtns}>
                  <IconButton iconProps={{ iconName: 'Up' }} title="Move section up" className={styles.iconBtn} disabled={si === 0} onClick={() => moveSection(si, -1)} />
                  <IconButton iconProps={{ iconName: 'Down' }} title="Move section down" className={styles.iconBtn} disabled={si === config.form.sections.length - 1} onClick={() => moveSection(si, 1)} />
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Remove section (fields move to the previous section)" className={styles.iconBtn} disabled={config.form.sections.length === 1} onClick={() => removeSection(si)} />
                </div>
              </div>
              <div className={styles.sectionBody}>
                <div className={styles.formRow} style={{ gridTemplateColumns: '1fr 1fr', marginBottom: 6 }}>
                  <TextField placeholder="Section description (optional)" value={section.description || ''} onChange={(_, v) => updateSection(si, s => { s.description = v || undefined; })} />
                  {section.collapsible && <Toggle label="Collapsed by default" inlineLabel checked={!!section.collapsedByDefault} onChange={(_, c) => updateSection(si, s => { s.collapsedByDefault = !!c; })} />}
                </div>
                {section.fields.length === 0 && <Text variant="small" className={styles.muted}>No fields yet – use the chips below to add some.</Text>}
                {section.fields.map((p, pi) => {
                  const f = fieldByName(config, p.field);
                  if (!f) return undefined;
                  return (
                    <div key={p.field} className={styles.placement}>
                      <div>
                        <div className={styles.placementName}>{f.displayName}{f.isSystem && <span className={styles.sysTag}>system</span>}</div>
                        <div className={styles.placementType}>{f.type}{f.required ? ' · required' : ''}</div>
                      </div>
                      <Dropdown selectedKey={p.width || 'half'} options={WIDTHS} disabled={(section.columns || 2) === 1}
                        onChange={(_, o) => updatePlacement(si, pi, x => { x.width = o?.key as FieldWidth; })} />
                      <TextField placeholder="Label override" value={p.label || ''} onChange={(_, v) => updatePlacement(si, pi, x => { x.label = v || undefined; })} />
                      <TextField placeholder={f.type === 'Text' || f.type === 'Note' || f.type === 'Number' || f.type === 'URL' ? 'Placeholder' : 'Help text'}
                        value={(f.type === 'Text' || f.type === 'Note' || f.type === 'Number' || f.type === 'URL' ? p.placeholder : p.helpText) || ''}
                        onChange={(_, v) => updatePlacement(si, pi, x => { if (f.type === 'Text' || f.type === 'Note' || f.type === 'Number' || f.type === 'URL') x.placeholder = v || undefined; else x.helpText = v || undefined; })} />
                      <Toggle title="Read-only" checked={!!p.readOnly || !!f.isSystem} disabled={!!f.isSystem} onText="RO" offText="RW"
                        onChange={(_, c) => updatePlacement(si, pi, x => { x.readOnly = !!c; })} styles={{ root: { marginBottom: 0 } }} />
                      <div className={styles.rowBtns}>
                        <IconButton iconProps={{ iconName: 'Up' }} className={styles.iconBtn} disabled={pi === 0} onClick={() => movePlacement(si, pi, -1)} />
                        <IconButton iconProps={{ iconName: 'Down' }} className={styles.iconBtn} disabled={pi === section.fields.length - 1} onClick={() => movePlacement(si, pi, 1)} />
                        <IconButton iconProps={{ iconName: 'Switch' }} className={styles.iconBtn} title="Move to another section" disabled={config.form.sections.length < 2}
                          menuProps={{ items: sectionOptions.filter(o => o.key !== section.id).map(o => ({ key: String(o.key), text: `Move to ${o.text}`, onClick: () => moveToSection(si, pi, String(o.key)) })) }} />
                        <IconButton iconProps={{ iconName: 'Hide3' }} className={styles.iconBtn} title="Hide from form" onClick={() => removePlacement(si, pi)} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}

          <div className={styles.addRow}>
            <DefaultButton text="Add section" iconProps={{ iconName: 'Add' }} onClick={addSection} />
          </div>

          {(unplaced.length > 0 || unplacedSystem.length > 0) && (
            <div className={styles.panel} style={{ marginTop: 12 }}>
              <div className={styles.panelTitle}>Not on the form – click to add</div>
              <div className={styles.chips}>
                {unplaced.map(f => <button key={f.internalName} type="button" className={styles.chip} onClick={() => place(f.internalName)}>＋ {f.displayName}</button>)}
                {unplacedSystem.map(f => <button key={f.internalName} type="button" className={styles.chip} onClick={() => place(f.internalName)} title="Read-only system field">＋ {f.displayName}</button>)}
              </div>
            </div>
          )}
        </div>

        <div className={styles.previewFrame}>
          <div className={styles.previewLabel}>
            <span>Live preview</span>
            <ChoiceGroup selectedKey={previewMode} styles={{ flexContainer: { display: 'flex', gap: 12 } }}
              options={[{ key: 'add', text: 'New' }, { key: 'edit', text: 'Edit' }, { key: 'view', text: 'View' }]}
              onChange={(_, o) => setPreviewMode(o?.key as 'add' | 'edit' | 'view')} />
          </div>
          <ItemForm
            key={`${previewMode}-${config.form.sections.length}`}
            inline
            config={config}
            mode={previewMode}
            item={previewMode === 'add' ? undefined : sample}
            onSave={async () => { window.alert('Preview only – nothing is saved from the designer.'); }}
            onDismiss={() => undefined}
            onEdit={() => setPreviewMode('edit')}
          />
        </div>
      </div>
    </div>
  );
};

export default FormStep;
