import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  TextField, Dropdown, IDropdownOption, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Spinner, SpinnerSize, Icon, Link, Text, Toggle
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { IRhinoConfig } from '../../../../shared/types/RhinoConfig';
import { toListInternalName, pluralise, buildDefaultForm, buildDefaultGrid } from '../../../../shared/config/helpers';
import { IValidationIssue, hasErrors } from '../../../../shared/config/validate';
import { ListSetupService, IListSummary, IProvisionReport } from '../../../../shared/services/ListSetupService';
import { normalizeConfig } from '../../../../shared/config/normalize';
import { withConfig, StepKey } from '../DesignerApp';
import styles from '../DesignerApp.module.scss';

export interface IListStepProps {
  sp: SPFI;
  config: IRhinoConfig;
  onChange: (next: IRhinoConfig) => void;
  setup: ListSetupService;
  lists: IListSummary[];
  onListsChanged: () => Promise<void>;
  webUrl: string;
  issues: IValidationIssue[];
  onGoToStep: (step: StepKey) => void;
}

const ACCENTS = ['#0f6cbd', '#107c10', '#8e3ba3', '#ca5010', '#038387', '#c239b3', '#5c2e91', '#4f6bed', '#a4262c', '#323130'];

export const ListStep: React.FC<IListStepProps> = ({ config, onChange, setup, lists, onListsChanged, webUrl, issues, onGoToStep }) => {
  const [mode, setMode] = useState<'new' | 'existing'>(() => lists.some(l => l.internalName === config.list.internalName) ? 'existing' : 'new');
  const [internalLocked, setInternalLocked] = useState<boolean>(!!config.list.internalName && config.list.internalName !== toListInternalName(config.list.title));
  const [provisioning, setProvisioning] = useState(false);
  const [report, setReport] = useState<IProvisionReport | undefined>();
  const [provisionError, setProvisionError] = useState<string | undefined>();
  const [exists, setExists] = useState<boolean | undefined>();
  const [importing, setImporting] = useState(false);

  const listIssues = issues.filter(i => i.path.indexOf('list.') === 0);
  const blocked = hasErrors(issues);

  // Lists load asynchronously – once they arrive, flip to "existing" if the design already points at one.
  useEffect(() => {
    if (config.list.internalName && lists.some(l => l.internalName === config.list.internalName)) setMode('existing');
  }, [lists.length]);

  useEffect(() => {
    let cancelled = false;
    setExists(undefined);
    if (!config.list.title) return;
    setup.listExists(config.list).then(e => { if (!cancelled) setExists(e); }).catch(() => { if (!cancelled) setExists(false); });
    return () => { cancelled = true; };
  }, [setup, config.list.title, config.list.internalName]);

  const setTitle = (title: string): void => {
    onChange(withConfig(config, d => {
      const wasSameAsTitle = !d.branding.title || d.branding.title === d.list.title;
      d.list.title = title;
      if (!internalLocked) d.list.internalName = toListInternalName(title);
      if (wasSameAsTitle) d.branding.title = title;
    }));
  };

  const pickExisting = (id: string): void => {
    const l = lists.filter(x => x.id === id)[0];
    if (!l) return;
    setInternalLocked(true);
    onChange(withConfig(config, d => {
      d.list = { title: l.title, internalName: l.internalName, description: l.description };
      if (!d.branding.title || d.branding.title === config.list.title) d.branding.title = l.title;
      if (!d.branding.subtitle) d.branding.subtitle = l.description;
    }));
    setReport(undefined);
  };

  const importColumns = async (): Promise<void> => {
    setImporting(true);
    try {
      const fields = await setup.importFields(config.list);
      const next = withConfig(config, d => {
        d.fields = fields;
        d.form = buildDefaultForm(fields);
        d.grid = buildDefaultGrid(fields);
      });
      onChange(normalizeConfig(next));
      onGoToStep('fields');
    } catch (e) {
      setProvisionError(`Could not import columns: ${(e as Error).message}`);
    } finally {
      setImporting(false);
    }
  };

  const provision = async (): Promise<void> => {
    setProvisioning(true);
    setProvisionError(undefined);
    setReport(undefined);
    try {
      const r = await setup.provisionList(config);
      setReport(r);
      setExists(true);
      await onListsChanged();
    } catch (e) {
      setProvisionError((e as Error).message);
    } finally {
      setProvisioning(false);
    }
  };

  const listOptions: IDropdownOption[] = lists.map(l => ({ key: l.id, text: `${l.title}  ·  /Lists/${l.internalName}  (${l.itemCount})` }));
  const selectedList = lists.filter(l => l.internalName === config.list.internalName || l.title === config.list.title)[0];
  const listUrl = `${webUrl}/Lists/${config.list.internalName || toListInternalName(config.list.title)}`;

  return (
    <div>
      <div className={styles.stepHeader}>
        <h2>Choose or create the list</h2>
        <p>Pick an existing SharePoint list, or define a new one. New lists get a clean URL from the internal name you control here.</p>
      </div>

      <div className={styles.modeCards}>
        <button type="button" className={`${styles.modeCard} ${mode === 'new' ? styles.modeCardActive : ''}`} onClick={() => setMode('new')}>
          <Icon iconName="AddTo" />
          <div><div className={styles.modeTitle}>New list</div><div className={styles.modeSub}>Define the name and columns, then provision it in one click.</div></div>
        </button>
        <button type="button" className={`${styles.modeCard} ${mode === 'existing' ? styles.modeCardActive : ''}`} onClick={() => setMode('existing')}>
          <Icon iconName="BulletedList" />
          <div><div className={styles.modeTitle}>Existing list</div><div className={styles.modeSub}>Bind to a list on this site. Import its columns, then add any that are missing.</div></div>
        </button>
      </div>

      <div className={styles.split}>
        <div className={styles.panel}>
          <div className={styles.panelTitle}>List</div>
          {mode === 'existing' && (
            <div className={styles.formRow} style={{ gridTemplateColumns: '1fr' }}>
              <Dropdown label="SharePoint list" options={listOptions} selectedKey={selectedList?.id} placeholder={lists.length ? 'Select a list' : 'No custom lists found on this site'}
                onChange={(_, o) => pickExisting(String(o?.key))} />
              <div className={styles.addRow}>
                <DefaultButton text={importing ? 'Importing…' : 'Import columns from this list'} iconProps={{ iconName: 'Download' }} disabled={!selectedList || importing}
                  onClick={() => { void importColumns(); }} />
                <Text variant="small" className={styles.muted}>Replaces the current fields, form and grid with the list&rsquo;s columns.</Text>
              </div>
            </div>
          )}
          <div className={styles.formRow}>
            <TextField label="List title" required value={config.list.title} onChange={(_, v) => setTitle(v || '')}
              errorMessage={listIssues.filter(i => i.path === 'list.title')[0]?.message} placeholder="e.g. Project Tracker" />
            <div>
              <TextField
                label="Internal name (URL)"
                value={config.list.internalName || ''}
                disabled={mode === 'existing' && !!selectedList}
                onChange={(_, v) => { setInternalLocked(true); onChange(withConfig(config, d => { d.list.internalName = (v || '').replace(/[^A-Za-z0-9_]/g, ''); })); }}
                errorMessage={listIssues.filter(i => i.path === 'list.internalName')[0]?.message}
                description={`/Lists/${config.list.internalName || toListInternalName(config.list.title) || '…'}`}
                prefix="Lists/"
              />
              {internalLocked && mode === 'new' && (
                <Link onClick={() => { setInternalLocked(false); onChange(withConfig(config, d => { d.list.internalName = toListInternalName(d.list.title); })); }} style={{ fontSize: 12 }}>
                  Reset to match title
                </Link>
              )}
            </div>
          </div>
          <TextField label="Description" multiline rows={2} value={config.list.description || ''}
            onChange={(_, v) => onChange(withConfig(config, d => { d.list.description = v || ''; }))} />

          {config.list.title && (
            <div style={{ marginTop: 10, fontSize: 12, color: '#616161', display: 'flex', alignItems: 'center', gap: 6 }}>
              {exists === undefined ? <Spinner size={SpinnerSize.xSmall} /> : <Icon iconName={exists ? 'CheckMark' : 'Info'} style={{ color: exists ? '#107c10' : '#8a8a8a' }} />}
              {exists === undefined ? 'Checking site…' : exists ? <span>List exists: <Link href={listUrl} target="_blank">{listUrl.replace(/^https?:\/\/[^/]+/, '')}</Link></span> : 'Not provisioned yet on this site.'}
            </div>
          )}

          <div className={styles.provisionBar}>
            <PrimaryButton
              text={provisioning ? 'Provisioning…' : exists ? 'Update list columns' : 'Provision list'}
              iconProps={{ iconName: exists ? 'Sync' : 'Database' }}
              onClick={() => { void provision(); }}
              disabled={provisioning || blocked || !config.list.title}
            />
            {provisioning && <Spinner size={SpinnerSize.small} />}
            <Text variant="small" className={styles.muted}>
              {exists
                ? 'Adds missing columns and syncs titles, descriptions, required flags and choices. Never deletes.'
                : `Creates /Lists/${config.list.internalName || toListInternalName(config.list.title) || '…'} with ${config.fields.filter(f => !f.isSystem).length} column${config.fields.length === 1 ? '' : 's'}.`}
            </Text>
            {blocked && <Text variant="small" style={{ color: '#a4262c' }}>Fix the validation errors first.</Text>}
          </div>

          {provisionError && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setProvisionError(undefined)} className={styles.notice}>{provisionError}</MessageBar>}
          {report && (
            <div className={styles.report}>
              <div className={styles.reportRow}><Icon iconName="Completed" style={{ color: '#107c10' }} /> <strong>{report.listCreated ? 'List created' : 'List updated'}</strong> <Link href={`${webUrl.replace(/\/$/, '').replace(/^(https?:\/\/[^/]+).*$/, '$1')}${report.listUrl}`} target="_blank">{report.listUrl}</Link></div>
              {report.created.length > 0 && <div>Created columns: <span className={styles.mono}>{report.created.join(', ')}</span></div>}
              {report.updated.length > 0 && <div>Updated columns: <span className={styles.mono}>{report.updated.join(', ')}</span></div>}
              {report.skipped.length > 0 && <div style={{ color: '#616161' }}>Unchanged: <span className={styles.mono}>{report.skipped.join(', ')}</span></div>}
              {report.warnings.length > 0 && <ul>{report.warnings.map((w, i) => <li key={i} style={{ color: '#7a4b00' }}>{w}</li>)}</ul>}
            </div>
          )}
        </div>

        <div className={styles.panel}>
          <div className={styles.panelTitle}>Branding</div>
          <div className={styles.formRow}>
            <TextField label="Header title" value={config.branding.title || ''} placeholder={config.list.title}
              onChange={(_, v) => onChange(withConfig(config, d => { d.branding.title = v || ''; }))} />
            <TextField label="Icon" value={config.branding.icon || ''} placeholder="🦏 or a Fluent icon name"
              description="An emoji, or a Fluent UI icon name such as ClipboardList"
              onChange={(_, v) => onChange(withConfig(config, d => { d.branding.icon = v || ''; }))} />
          </div>
          <TextField label="Subtitle" value={config.branding.subtitle || ''}
            onChange={(_, v) => onChange(withConfig(config, d => { d.branding.subtitle = v || ''; }))} />
          <div className={styles.formRow} style={{ marginTop: 12 }}>
            <TextField label="Item label (singular)" value={config.branding.itemLabel || ''} placeholder="item"
              onChange={(_, v) => onChange(withConfig(config, d => { d.branding.itemLabel = v || ''; d.branding.itemLabelPlural = pluralise(v || 'item'); }))} />
            <TextField label="Item label (plural)" value={config.branding.itemLabelPlural || ''} placeholder="items"
              onChange={(_, v) => onChange(withConfig(config, d => { d.branding.itemLabelPlural = v || ''; }))} />
          </div>
          <div style={{ marginTop: 12 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 6 }}>Accent colour</Text>
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
              {ACCENTS.map(c => (
                <button key={c} type="button" title={c} onClick={() => onChange(withConfig(config, d => { d.branding.accent = c; }))}
                  style={{ width: 26, height: 26, borderRadius: '50%', background: c, border: config.branding.accent === c ? '3px solid #fff' : '2px solid transparent', boxShadow: config.branding.accent === c ? `0 0 0 2px ${c}` : 'none', cursor: 'pointer' }} />
              ))}
              <TextField value={config.branding.accent || ''} placeholder="#0f6cbd" styles={{ root: { width: 110 } }}
                onChange={(_, v) => onChange(withConfig(config, d => { d.branding.accent = v || undefined; }))} />
            </div>
          </div>
          <div style={{ marginTop: 14 }}>
            <Toggle label="Show created / modified info on the form" checked={config.form.showSystemInfo !== false}
              onChange={(_, c) => onChange(withConfig(config, d => { d.form.showSystemInfo = !!c; }))} />
          </div>

          <div className={styles.docsBox} style={{ marginTop: 16 }}>
            <strong>How it fits together</strong><br />
            1. Define the list and fields here, then <em>Provision list</em>.<br />
            2. Lay out the form and configure the grid.<br />
            3. <em>Save to site</em> writes <code>SiteAssets/RhinoForms/{config.list.internalName || 'name'}.json</code>.<br />
            4. Add a <strong>Rhino List</strong> web part to any page and pick that file.
          </div>
        </div>
      </div>
    </div>
  );
};

export default ListStep;
