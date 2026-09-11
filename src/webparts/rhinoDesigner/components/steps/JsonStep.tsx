import * as React from 'react';
import { useState, useEffect } from 'react';
import { PrimaryButton, DefaultButton, MessageBar, MessageBarType, Text, Icon } from '@fluentui/react';
import { IRhinoConfig } from '../../../../shared/types/RhinoConfig';
import { parseConfig, serializeConfig } from '../../../../shared/config/normalize';
import { IValidationIssue } from '../../../../shared/config/validate';
import styles from '../DesignerApp.module.scss';

export interface IJsonStepProps {
  config: IRhinoConfig;
  onChange: (next: IRhinoConfig) => void;
  issues: IValidationIssue[];
  fileUrl?: string;
  webUrl: string;
}

export const JsonStep: React.FC<IJsonStepProps> = ({ config, onChange, issues, fileUrl, webUrl }) => {
  const [text, setText] = useState(() => serializeConfig(config));
  const [parseError, setParseError] = useState<string | undefined>();
  const [edited, setEdited] = useState(false);

  useEffect(() => {
    if (!edited) setText(serializeConfig(config));
  }, [config, edited]);

  const apply = (): void => {
    try {
      onChange(parseConfig(text));
      setParseError(undefined);
      setEdited(false);
    } catch (e) {
      setParseError((e as Error).message);
    }
  };

  const reset = (): void => {
    setText(serializeConfig(config));
    setParseError(undefined);
    setEdited(false);
  };

  const origin = webUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');

  return (
    <div>
      <div className={styles.stepHeader}>
        <h2>The configuration file</h2>
        <p>This JSON is the single source of truth. Edit it here, in the saved file, or in source control – the Rhino List web part re-reads it on load.</p>
      </div>

      <div className={styles.split} style={{ gridTemplateColumns: 'minmax(420px, 7fr) minmax(300px, 4fr)' }}>
        <div>
          <textarea
            className={styles.jsonArea}
            value={text}
            spellCheck={false}
            onChange={e => { setText(e.target.value); setEdited(true); }}
          />
          <div className={styles.addRow}>
            <PrimaryButton text="Apply changes" iconProps={{ iconName: 'CheckMark' }} onClick={apply} disabled={!edited} />
            <DefaultButton text="Revert" iconProps={{ iconName: 'Undo' }} onClick={reset} disabled={!edited} />
            {edited && <Text variant="small" style={{ color: '#7a4b00' }}>Unapplied edits</Text>}
          </div>
          {parseError && <MessageBar messageBarType={MessageBarType.error} className={styles.notice}>{parseError}</MessageBar>}
        </div>

        <div>
          {fileUrl && (
            <div className={styles.panel} style={{ marginBottom: 12 }}>
              <div className={styles.panelTitle}>Saved file</div>
              <Text variant="small"><Icon iconName="Document" /> <a href={`${origin}${fileUrl}`} target="_blank" rel="noopener noreferrer">{fileUrl}</a></Text>
            </div>
          )}

          <div className={styles.panel} style={{ marginBottom: 12 }}>
            <div className={styles.panelTitle}>Validation</div>
            {issues.length === 0
              ? <Text variant="small" style={{ color: '#0e6b0e' }}><Icon iconName="Completed" /> No issues</Text>
              : issues.map((i, idx) => (
                <div key={idx} className={`${styles.issue} ${i.level === 'error' ? styles.issueError : styles.issueWarn}`} style={{ marginBottom: 4 }}>
                  <Icon iconName={i.level === 'error' ? 'ErrorBadge' : 'Warning'} /> <code>{i.path}</code> {i.message}
                </div>
              ))}
          </div>

          <div className={styles.docsBox}>
            <strong>Schema cheat-sheet</strong>
            <ul style={{ margin: '6px 0 0 16px', padding: 0 }}>
              <li><code>list</code> – <code>title</code>, <code>internalName</code> (URL segment), <code>description</code></li>
              <li><code>branding</code> – <code>title</code>, <code>subtitle</code>, <code>icon</code>, <code>accent</code>, <code>itemLabel</code>, <code>itemLabelPlural</code></li>
              <li><code>fields[]</code> – <code>internalName</code>, <code>displayName</code>, <code>type</code> (Text, Note, Choice, MultiChoice, DateTime, Number, Boolean, User, URL), <code>required</code>, <code>choices</code>, <code>defaultValue</code>, <code>description</code>, <code>maxLength</code>, <code>numberFormat</code>, <code>decimals</code>, <code>min</code>, <code>max</code>, <code>dateOnly</code>, <code>allowMultiple</code></li>
              <li><code>form</code> – <code>surface</code> (panel | dialog), <code>size</code>, <code>addTitle</code>, <code>editTitle</code>, <code>showSystemInfo</code>, <code>sections[]</code> of <code>{'{ id, title, columns, collapsible, fields: [{ field, width, label, placeholder, helpText, readOnly }] }'}</code></li>
              <li><code>grid</code> – <code>columns[]</code> of <code>{'{ field, label, width, sortable, filterable, render, align, badgeTones }'}</code>, <code>defaultSort</code>, <code>pageSize</code>, <code>density</code>, <code>rowClick</code>, <code>allowSearch / allowFilter / allowSort / allowExport / allowAdd / allowEdit / allowDelete / allowColumnResize</code>, <code>showCount</code>, <code>emptyMessage</code></li>
            </ul>
            <div style={{ marginTop: 8 }}>Anything you leave out is filled with sensible defaults. v1 files (<code>listName</code> + <code>fields</code>) still load.</div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default JsonStep;
