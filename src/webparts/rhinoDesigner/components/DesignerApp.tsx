import * as React from 'react';
import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  Pivot, PivotItem, PrimaryButton, DefaultButton, CommandBarButton, IContextualMenuProps, MessageBar, MessageBarType, Dialog, DialogType,
  DialogFooter, TextField, Dropdown, IDropdownOption, Spinner, SpinnerSize, Icon, Text
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { IRhinoConfig } from '../../../shared/types/RhinoConfig';
import { createEmptyConfig, sampleConfig, cloneConfig, configFileName, buildDefaultForm, buildDefaultGrid } from '../../../shared/config/helpers';
import { parseConfig, serializeConfig, normalizeConfig } from '../../../shared/config/normalize';
import { validateConfig, hasErrors, IValidationIssue } from '../../../shared/config/validate';
import { ConfigStore, IStoredConfig } from '../../../shared/services/ConfigStore';
import { ListSetupService, IListSummary } from '../../../shared/services/ListSetupService';
import ListStep from './steps/ListStep';
import FieldsStep from './steps/FieldsStep';
import FormStep from './steps/FormStep';
import GridStep from './steps/GridStep';
import JsonStep from './steps/JsonStep';
import styles from './DesignerApp.module.scss';

export interface IDesignerAppProps {
  sp: SPFI;
  webUrl: string;
  webServerRelativeUrl: string;
  draftKey: string;
}

export type StepKey = 'list' | 'fields' | 'form' | 'grid' | 'json';

interface IDraft {
  config: IRhinoConfig;
  fileUrl?: string;
  savedAt?: string;
}

function readDraft(key: string): IDraft | undefined {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return undefined;
    const d = JSON.parse(raw) as IDraft;
    return { ...d, config: normalizeConfig(d.config) };
  } catch {
    return undefined;
  }
}

function writeDraft(key: string, draft: IDraft): void {
  try { window.localStorage.setItem(key, JSON.stringify(draft)); } catch { /* quota / private mode */ }
}

export const DesignerApp: React.FC<IDesignerAppProps> = ({ sp, webUrl, draftKey }) => {
  const store = useMemo(() => new ConfigStore(sp), [sp]);
  const setup = useMemo(() => new ListSetupService(sp), [sp]);

  const [config, setConfig] = useState<IRhinoConfig>(() => readDraft(draftKey)?.config || createEmptyConfig());
  const [fileUrl, setFileUrl] = useState<string | undefined>(() => readDraft(draftKey)?.fileUrl);
  const [dirty, setDirty] = useState(false);
  const [step, setStep] = useState<StepKey>('list');
  const [files, setFiles] = useState<IStoredConfig[]>([]);
  const [lists, setLists] = useState<IListSummary[]>([]);
  const [busy, setBusy] = useState<string | undefined>();
  const [notice, setNotice] = useState<{ type: MessageBarType; text: string } | undefined>();
  const [openDialog, setOpenDialog] = useState(false);
  const [saveDialog, setSaveDialog] = useState(false);
  const [saveName, setSaveName] = useState('');
  const [pickedFile, setPickedFile] = useState<string | undefined>();
  const [showIssues, setShowIssues] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  const issues: IValidationIssue[] = useMemo(() => validateConfig(config), [config]);
  const blocked = hasErrors(issues);

  const refreshLists = useCallback(async () => {
    try { setLists(await setup.getLists()); } catch { setLists([]); }
  }, [setup]);
  const refreshFiles = useCallback(async () => {
    try { setFiles(await store.list()); } catch { setFiles([]); }
  }, [store]);

  useEffect(() => { void refreshLists(); void refreshFiles(); }, [refreshLists, refreshFiles]);

  // Draft autosave (debounced).
  useEffect(() => {
    const t = setTimeout(() => writeDraft(draftKey, { config, fileUrl, savedAt: new Date().toISOString() }), 400);
    return () => clearTimeout(t);
  }, [config, fileUrl, draftKey]);

  useEffect(() => {
    if (!notice) return;
    const t = setTimeout(() => setNotice(undefined), 6000);
    return () => clearTimeout(t);
  }, [notice]);

  const update = useCallback((next: IRhinoConfig): void => {
    setConfig(next);
    setDirty(true);
  }, []);

  const replaceConfig = (next: IRhinoConfig, url?: string, message?: string): void => {
    setConfig(next);
    setFileUrl(url);
    setDirty(false);
    if (message) setNotice({ type: MessageBarType.success, text: message });
  };

  const confirmDiscard = (): boolean => !dirty || window.confirm('You have unsaved changes. Replace the current design?');

  // ---- load actions ------------------------------------------------------
  const loadFromSite = async (url: string): Promise<void> => {
    if (!confirmDiscard()) return;
    setBusy('Loading config…');
    try {
      const cfg = await store.load(url);
      replaceConfig(cfg, url, `Loaded ${url.split('/').pop()}`);
      setOpenDialog(false);
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Could not load: ${(e as Error).message}` });
    } finally {
      setBusy(undefined);
    }
  };

  const loadFromFile = (file: File): void => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const cfg = parseConfig(String(ev.target?.result || ''));
        if (!confirmDiscard()) return;
        replaceConfig(cfg, undefined, `Loaded ${file.name}`);
        setOpenDialog(false);
      } catch (e) {
        setNotice({ type: MessageBarType.error, text: `Invalid config file: ${(e as Error).message}` });
      }
    };
    reader.readAsText(file);
  };

  const startFromList = async (list: IListSummary): Promise<void> => {
    setBusy(`Reading columns from "${list.title}"…`);
    try {
      const fields = await setup.importFields({ title: list.title, internalName: list.internalName });
      const cfg = createEmptyConfig();
      cfg.list = { title: list.title, internalName: list.internalName, description: list.description };
      cfg.branding = { ...cfg.branding, title: list.title, subtitle: list.description };
      cfg.fields = fields.length ? fields : cfg.fields;
      cfg.form = buildDefaultForm(cfg.fields);
      cfg.grid = buildDefaultGrid(cfg.fields);
      replaceConfig(normalizeConfig(cfg), undefined, `Imported ${fields.length} columns from "${list.title}"`);
      setOpenDialog(false);
      setStep('fields');
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Could not read list: ${(e as Error).message}` });
    } finally {
      setBusy(undefined);
    }
  };

  const newBlank = (): void => { if (confirmDiscard()) { replaceConfig(createEmptyConfig(), undefined, 'Started a blank design'); setOpenDialog(false); setStep('list'); } };
  const newSample = (): void => { if (confirmDiscard()) { replaceConfig(sampleConfig(), undefined, 'Loaded the sample "Project Tracker" design'); setOpenDialog(false); setStep('list'); } };

  // ---- save actions ------------------------------------------------------
  const openSave = (): void => {
    setSaveName((fileUrl ? fileUrl.split('/').pop() : configFileName(config)) || 'rhino-config.json');
    setSaveDialog(true);
  };

  const saveToSite = async (): Promise<void> => {
    setBusy('Saving to Site Assets…');
    try {
      const saved = await store.save(saveName, config);
      setFileUrl(saved.serverRelativeUrl);
      setDirty(false);
      setSaveDialog(false);
      setNotice({ type: MessageBarType.success, text: `Saved ${saved.name} to Site Assets/RhinoForms. Rhino List web parts pointing at it will pick it up on next load.` });
      void refreshFiles();
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Save failed: ${(e as Error).message}` });
    } finally {
      setBusy(undefined);
    }
  };

  const download = (): void => {
    const blob = new Blob([serializeConfig(config)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = configFileName(config);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const copyJson = async (): Promise<void> => {
    try {
      await navigator.clipboard.writeText(serializeConfig(config));
      setNotice({ type: MessageBarType.success, text: 'JSON copied to clipboard' });
    } catch {
      setNotice({ type: MessageBarType.warning, text: 'Clipboard blocked – use the JSON tab to copy manually.' });
    }
  };

  const loadMenu: IContextualMenuProps = {
    items: [
      { key: 'site', text: 'From Site Assets…', iconProps: { iconName: 'SharepointLogo' }, onClick: () => { setPickedFile(undefined); setOpenDialog(true); } },
      { key: 'file', text: 'From a JSON file…', iconProps: { iconName: 'OpenFile' }, onClick: () => fileInput.current?.click() },
      { key: 'div', itemType: 1 },
      { key: 'blank', text: 'New blank design', iconProps: { iconName: 'Page' }, onClick: newBlank },
      { key: 'sample', text: 'New from sample (Project Tracker)', iconProps: { iconName: 'Lightbulb' }, onClick: newSample }
    ]
  };

  const stepStatus = (key: StepKey): string | undefined => {
    const prefix = key === 'list' ? 'list.' : key === 'fields' ? 'fields' : key === 'form' ? 'form.' : key === 'grid' ? 'grid.' : '';
    if (!prefix) return undefined;
    return issues.some(i => i.level === 'error' && i.path.indexOf(prefix) === 0) ? '●' : undefined;
  };

  const renderStep = (): React.ReactNode => {
    switch (step) {
      case 'list': return <ListStep sp={sp} config={config} onChange={update} setup={setup} lists={lists} onListsChanged={refreshLists} webUrl={webUrl} issues={issues} onGoToStep={setStep} />;
      case 'fields': return <FieldsStep config={config} onChange={update} issues={issues} />;
      case 'form': return <FormStep config={config} onChange={update} />;
      case 'grid': return <GridStep config={config} onChange={update} />;
      case 'json': return <JsonStep config={config} onChange={update} issues={issues} fileUrl={fileUrl} webUrl={webUrl} />;
      default: return undefined;
    }
  };

  const pivotItem = (key: StepKey, text: string, icon: string): JSX.Element => (
    <PivotItem key={key} itemKey={key} headerText={text} itemIcon={icon} onRenderItemLink={(link, defaultRender) => (
      <span className={styles.pivotLink}>{defaultRender ? defaultRender(link) : text}{stepStatus(key) && <span className={styles.pivotError} title="This step has errors">{stepStatus(key)}</span>}</span>
    )} />
  );

  return (
    <div className={styles.app}>
      <div className={styles.topBar}>
        <div className={styles.brand}>
          <span className={styles.brandIcon}>🦏</span>
          <div>
            <div className={styles.brandTitle}>Rhino List Designer</div>
            <div className={styles.brandSub}>
              {config.list.title ? <strong>{config.list.title}</strong> : <em>Untitled design</em>}
              {fileUrl && <span className={styles.fileTag}><Icon iconName="Document" /> {fileUrl.split('/').pop()}</span>}
              {dirty && <span className={styles.dirtyTag}>unsaved changes</span>}
            </div>
          </div>
        </div>
        <div className={styles.actions}>
          <CommandBarButton text="Open" iconProps={{ iconName: 'FolderOpen' }} menuProps={loadMenu} className={styles.cmdBtn} />
          <CommandBarButton text="Download" iconProps={{ iconName: 'Download' }} onClick={download} className={styles.cmdBtn} />
          <CommandBarButton text="Copy JSON" iconProps={{ iconName: 'Copy' }} onClick={() => { void copyJson(); }} className={styles.cmdBtn} />
          <button type="button" className={`${styles.issuesBtn} ${blocked ? styles.issuesBad : issues.length ? styles.issuesWarn : styles.issuesOk}`} onClick={() => setShowIssues(!showIssues)}>
            <Icon iconName={blocked ? 'ErrorBadge' : issues.length ? 'Warning' : 'Completed'} />
            {blocked ? `${issues.filter(i => i.level === 'error').length} error${issues.filter(i => i.level === 'error').length === 1 ? '' : 's'}` : issues.length ? `${issues.length} warning${issues.length === 1 ? '' : 's'}` : 'Valid'}
          </button>
          <PrimaryButton text="Save to site" iconProps={{ iconName: 'Save' }} onClick={openSave} disabled={blocked || !!busy} />
        </div>
      </div>

      {showIssues && issues.length > 0 && (
        <div className={styles.issuesPanel}>
          {issues.map((i, idx) => (
            <div key={idx} className={`${styles.issue} ${i.level === 'error' ? styles.issueError : styles.issueWarn}`}>
              <Icon iconName={i.level === 'error' ? 'ErrorBadge' : 'Warning'} /> <code>{i.path}</code> {i.message}
            </div>
          ))}
        </div>
      )}

      {notice && <MessageBar messageBarType={notice.type} onDismiss={() => setNotice(undefined)} className={styles.notice}>{notice.text}</MessageBar>}
      {busy && <div className={styles.busy}><Spinner size={SpinnerSize.small} label={busy} labelPosition="right" /></div>}

      <Pivot selectedKey={step} onLinkClick={item => item && setStep(item.props.itemKey as StepKey)} className={styles.pivot} headersOnly>
        {pivotItem('list', '1. List', 'BulletedList')}
        {pivotItem('fields', '2. Fields', 'FieldChanged')}
        {pivotItem('form', '3. Form layout', 'PageEdit')}
        {pivotItem('grid', '4. Grid', 'Table')}
        {pivotItem('json', '5. JSON', 'Code')}
      </Pivot>

      <div className={styles.stepBody}>{renderStep()}</div>

      <div className={styles.stepNav}>
        <DefaultButton text="Back" iconProps={{ iconName: 'ChevronLeft' }} disabled={step === 'list'}
          onClick={() => setStep((['list', 'fields', 'form', 'grid', 'json'] as StepKey[])[Math.max(0, ['list', 'fields', 'form', 'grid', 'json'].indexOf(step) - 1)])} />
        <div className={styles.stepNavHint}><Text variant="small">Drafts are kept in this browser automatically. Use <strong>Save to site</strong> to publish for Rhino List web parts.</Text></div>
        {step !== 'json'
          ? <PrimaryButton text="Next" iconProps={{ iconName: 'ChevronRight' }} onClick={() => setStep((['list', 'fields', 'form', 'grid', 'json'] as StepKey[])[['list', 'fields', 'form', 'grid', 'json'].indexOf(step) + 1])} />
          : <PrimaryButton text="Save to site" iconProps={{ iconName: 'Save' }} onClick={openSave} disabled={blocked || !!busy} />}
      </div>

      <input ref={fileInput} type="file" accept=".json,application/json" style={{ display: 'none' }}
        onChange={e => { const f = e.target.files?.[0]; if (f) loadFromFile(f); e.target.value = ''; }} />

      <Dialog hidden={!openDialog} onDismiss={() => setOpenDialog(false)} minWidth={520}
        dialogContentProps={{ type: DialogType.largeHeader, title: 'Open a design' }}>
        <div className={styles.dialogSection}>
          <Text className={styles.dialogLabel}>Saved configs (Site Assets/RhinoForms)</Text>
          {files.length === 0
            ? <Text variant="small" className={styles.muted}>No saved configs on this site yet.</Text>
            : <Dropdown options={files.map<IDropdownOption>(f => ({ key: f.serverRelativeUrl, text: f.name }))} selectedKey={pickedFile} placeholder="Choose a config file" onChange={(_, o) => setPickedFile(o?.key as string)} />}
          <PrimaryButton text="Open config" disabled={!pickedFile} onClick={() => pickedFile && void loadFromSite(pickedFile)} className={styles.dialogBtn} />
        </div>
        <div className={styles.dialogSection}>
          <Text className={styles.dialogLabel}>Start from an existing list</Text>
          <Text variant="small" className={styles.muted}>Reads the list&rsquo;s columns and builds a default form and grid you can then refine.</Text>
          <Dropdown options={lists.map<IDropdownOption>(l => ({ key: l.id, text: `${l.title}  (${l.itemCount} items)` }))} placeholder="Choose a list"
            onChange={(_, o) => { const l = lists.filter(x => x.id === o?.key)[0]; if (l) void startFromList(l); }} />
        </div>
        <DialogFooter>
          <DefaultButton text="Close" onClick={() => setOpenDialog(false)} />
        </DialogFooter>
      </Dialog>

      <Dialog hidden={!saveDialog} onDismiss={() => setSaveDialog(false)} minWidth={480}
        dialogContentProps={{ type: DialogType.largeHeader, title: 'Save to Site Assets', subText: 'The file is written to Site Assets/RhinoForms. Rhino List web parts select it by name.' }}>
        <TextField label="File name" value={saveName} onChange={(_, v) => setSaveName(v || '')} suffix={/\.json$/i.test(saveName) ? undefined : '.json'} />
        {files.some(f => f.name.toLowerCase() === (/\.json$/i.test(saveName) ? saveName : `${saveName}.json`).toLowerCase()) && (
          <MessageBar messageBarType={MessageBarType.warning} className={styles.notice}>A file with this name exists and will be overwritten.</MessageBar>
        )}
        <DialogFooter>
          <PrimaryButton text="Save" onClick={() => { void saveToSite(); }} disabled={!saveName.trim() || !!busy} />
          <DefaultButton text="Cancel" onClick={() => setSaveDialog(false)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default DesignerApp;

/** Small helpers shared by the steps. */
export function moveItem<T>(arr: T[], from: number, to: number): T[] {
  const copy = arr.slice();
  if (to < 0 || to >= copy.length) return copy;
  const [it] = copy.splice(from, 1);
  copy.splice(to, 0, it);
  return copy;
}

export function withConfig(config: IRhinoConfig, mutate: (draft: IRhinoConfig) => void): IRhinoConfig {
  const draft = cloneConfig(config);
  mutate(draft);
  return draft;
}
