import * as React from 'react';
import { useState, useEffect } from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, PrimaryButton, DefaultButton, Icon, Text, Link } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IRhinoConfig } from '../../../shared/types/RhinoConfig';
import { parseConfig } from '../../../shared/config/normalize';
import { ConfigStore, IStoredConfig } from '../../../shared/services/ConfigStore';
import RhinoList from '../../../shared/components/RhinoList/RhinoList';
import type { ConfigSource } from '../RhinoGridWebPart';
import styles from './RhinoGridApp.module.scss';

export interface IRhinoGridAppProps {
  sp: SPFI;
  store: ConfigStore;
  configSource: ConfigSource;
  configFileUrl: string;
  inlineJson: string;
  hideHeader: boolean;
  displayMode: DisplayMode;
  designerUrl: string;
  onOpenPropertyPane: () => void;
  onPickConfigFile: (url: string) => void;
}

type LoadState =
  | { kind: 'idle' }
  | { kind: 'loading' }
  | { kind: 'ready'; config: IRhinoConfig; source: string }
  | { kind: 'error'; message: string };

const RhinoGridApp: React.FC<IRhinoGridAppProps> = (props) => {
  const { sp, store, configSource, configFileUrl, inlineJson, hideHeader, displayMode, designerUrl, onOpenPropertyPane, onPickConfigFile } = props;
  const [state, setState] = useState<LoadState>({ kind: 'idle' });
  const [files, setFiles] = useState<IStoredConfig[] | undefined>();
  const isEdit = displayMode === DisplayMode.Edit;

  useEffect(() => {
    let cancelled = false;
    const run = async (): Promise<void> => {
      if (configSource === 'inline') {
        if (!inlineJson.trim()) { setState({ kind: 'idle' }); return; }
        try {
          setState({ kind: 'ready', config: parseConfig(inlineJson), source: 'inline JSON' });
        } catch (e) {
          setState({ kind: 'error', message: (e as Error).message });
        }
        return;
      }
      if (!configFileUrl) { setState({ kind: 'idle' }); return; }
      setState({ kind: 'loading' });
      try {
        const config = await store.load(configFileUrl);
        if (!cancelled) setState({ kind: 'ready', config, source: configFileUrl.split('/').pop() || configFileUrl });
      } catch (e) {
        if (!cancelled) setState({ kind: 'error', message: `Could not load "${configFileUrl}": ${(e as Error).message}` });
      }
    };
    void run();
    return () => { cancelled = true; };
  }, [store, configSource, configFileUrl, inlineJson]);

  useEffect(() => {
    if (state.kind === 'idle' && isEdit && files === undefined) {
      store.list().then(setFiles).catch(() => setFiles([]));
    }
  }, [state.kind, isEdit, files, store]);

  if (state.kind === 'loading') {
    return <div className={styles.center}><Spinner size={SpinnerSize.large} label="Loading configuration…" /></div>;
  }

  if (state.kind === 'error') {
    return (
      <div className={styles.center}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          <strong>Rhino List could not read its configuration.</strong><br />{state.message}
          {isEdit && <><br /><Link onClick={onOpenPropertyPane}>Open the web part settings</Link> to pick another config.</>}
        </MessageBar>
      </div>
    );
  }

  if (state.kind === 'idle') {
    return (
      <div className={styles.center}>
        <div className={styles.setupCard}>
          <div className={styles.setupIcon}>🦏</div>
          <Text variant="xLarge" className={styles.setupTitle}>Rhino List needs a configuration</Text>
          <Text variant="medium" className={styles.setupSub}>
            Pick a config file created with the Rhino List Designer, or paste JSON in the web part settings.
          </Text>
          {isEdit ? (
            <>
              {files && files.length > 0 && (
                <div className={styles.fileList}>
                  {files.map(f => (
                    <button key={f.serverRelativeUrl} type="button" className={styles.fileBtn} onClick={() => onPickConfigFile(f.serverRelativeUrl)}>
                      <Icon iconName="Document" /> <span>{f.name.replace(/\.json$/i, '')}</span>
                    </button>
                  ))}
                </div>
              )}
              {files && files.length === 0 && (
                <MessageBar messageBarType={MessageBarType.info}>No config files yet. Add the Designer web part to a page to create one.</MessageBar>
              )}
              <div className={styles.setupActions}>
                <PrimaryButton text="Web part settings" iconProps={{ iconName: 'Settings' }} onClick={onOpenPropertyPane} />
                <DefaultButton text="Open Designer" iconProps={{ iconName: 'Design' }} href={designerUrl} target="_blank" />
              </div>
            </>
          ) : (
            <MessageBar messageBarType={MessageBarType.info}>Edit this page to configure the web part.</MessageBar>
          )}
        </div>
      </div>
    );
  }

  const headerExtra = isEdit ? (
    <Link onClick={onOpenPropertyPane} className={styles.sourceLink} title={`Config: ${state.source}`}>
      <Icon iconName="Settings" /> {state.source}
    </Link>
  ) : undefined;

  return <RhinoList sp={sp} config={state.config} hideHeader={hideHeader} headerExtra={headerExtra} />;
};

export default RhinoGridApp;
