import * as React from 'react';
import { useState, useRef } from 'react';
import {
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Icon,
  Text
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { IGridConfig } from '../../types/IFieldConfig';
import { ListSetupService } from '../../services/ListSetupService';
import styles from './ConfigSetup.module.scss';

export interface IConfigSetupProps {
  sp: SPFI;
  setupService: ListSetupService;
  existingConfig?: IGridConfig;
  onConfigSaved: (json: string, listName: string) => void;
}

export const ConfigSetup: React.FC<IConfigSetupProps> = ({ setupService, existingConfig, onConfigSaved }) => {
  const [parsedConfig, setParsedConfig] = useState<IGridConfig | undefined>(existingConfig);
  const [jsonError, setJsonError] = useState<string | undefined>();
  const [provisioning, setProvisioning] = useState(false);
  const [provisionError, setProvisionError] = useState<string | undefined>();
  const [provisionSuccess, setProvisionSuccess] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
      const content = ev.target?.result as string;
      parseJson(content);
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const parseJson = (content: string): void => {
    setJsonError(undefined);
    setProvisionError(undefined);
    setProvisionSuccess(false);
    try {
      const config = JSON.parse(content) as IGridConfig;
      if (!config.listName) throw new Error('"listName" is required in the configuration.');
      if (!Array.isArray(config.fields) || config.fields.length === 0) {
        throw new Error('"fields" array is required and must not be empty.');
      }
      for (const field of config.fields) {
        if (!field.internalName) throw new Error(`Each field must have an "internalName".`);
        if (!field.displayName) throw new Error(`Field "${field.internalName}" must have a "displayName".`);
        if (!field.type) throw new Error(`Field "${field.internalName}" must have a "type".`);
      }
      setParsedConfig(config);
    } catch (err) {
      setJsonError(`Invalid JSON configuration: ${err.message}`);
      setParsedConfig(undefined);
    }
  };

  const handleProvision = async (): Promise<void> => {
    if (!parsedConfig) return;
    setProvisioning(true);
    setProvisionError(undefined);
    setProvisionSuccess(false);
    try {
      await setupService.provisionList(parsedConfig);
      setProvisionSuccess(true);
      onConfigSaved(JSON.stringify(parsedConfig), parsedConfig.listName);
    } catch (err) {
      setProvisionError(`Failed to provision list: ${err.message}`);
    } finally {
      setProvisioning(false);
    }
  };

  const isReProvisioning = !!existingConfig;

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text variant="xLarge">{isReProvisioning ? 'Re-configure Rhino Grid' : 'Set Up Rhino Grid'}</Text>
        <br />
        <Text variant="medium" style={{ color: '#605e5c' }}>
          {isReProvisioning
            ? `Currently using list "${existingConfig.listName}". Upload a new config or re-provision the existing list.`
            : 'Upload a JSON configuration file to define your grid columns and provision the SharePoint list.'}
        </Text>
      </div>

      <div className={styles.uploadArea} onClick={() => fileInputRef.current?.click()}>
        <div className={styles.uploadIcon}>
          <Icon iconName="Upload" />
        </div>
        <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Click to upload JSON config</Text>
        <br />
        <Text variant="small" style={{ color: '#605e5c' }}>
          or drag and drop a .json file here
        </Text>
        <input
          ref={fileInputRef}
          type="file"
          accept=".json,application/json"
          className={styles.fileInput}
          onChange={handleFileSelect}
        />
      </div>

      {jsonError && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setJsonError(undefined)}>
          {jsonError}
        </MessageBar>
      )}

      {parsedConfig && (
        <>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
            Preview: {parsedConfig.fields.length} fields for list &ldquo;{parsedConfig.listName}&rdquo;
          </Text>
          <table className={styles.previewTable}>
            <thead>
              <tr>
                <th>Internal Name</th>
                <th>Display Name</th>
                <th>Type</th>
                <th>Options</th>
              </tr>
            </thead>
            <tbody>
              {parsedConfig.fields.map(field => (
                <tr key={field.internalName}>
                  <td><code>{field.internalName}</code></td>
                  <td>{field.displayName}</td>
                  <td>{field.type}</td>
                  <td>
                    {field.required && <span className={`${styles.badge} ${styles.required}`}>Required</span>}
                    {field.filterable && <span className={`${styles.badge} ${styles.filterable}`}>Filterable</span>}
                    {field.choices && field.choices.length > 0 && (
                      <span className={styles.badge}>{field.choices.join(', ')}</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className={styles.actions}>
            <PrimaryButton
              text={isReProvisioning ? 'Re-provision List & Save' : 'Provision List & Start'}
              iconProps={{ iconName: 'Database' }}
              onClick={handleProvision}
              disabled={provisioning}
            />
            {provisioning && <Spinner size={SpinnerSize.small} label="Provisioning list..." />}
          </div>

          <div className={styles.provisionStatus}>
            {provisionError && (
              <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setProvisionError(undefined)}>
                {provisionError}
              </MessageBar>
            )}
            {provisionSuccess && (
              <MessageBar messageBarType={MessageBarType.success}>
                List &ldquo;{parsedConfig.listName}&rdquo; provisioned successfully! The grid is now ready.
              </MessageBar>
            )}
          </div>
        </>
      )}

      {!parsedConfig && (
        <details style={{ marginTop: 16 }}>
          <summary style={{ cursor: 'pointer', color: '#0078d4' }}>View example JSON configuration</summary>
          <pre style={{ background: '#f3f2f1', padding: 16, borderRadius: 4, fontSize: 12, overflowX: 'auto', marginTop: 8 }}>
{`{
  "listName": "Project Tracker",
  "listDescription": "Tracks projects and assignments",
  "fields": [
    {
      "internalName": "Title",
      "displayName": "Project Name",
      "type": "Text",
      "required": true,
      "filterable": true,
      "sortable": true
    },
    {
      "internalName": "Status",
      "displayName": "Status",
      "type": "Choice",
      "choices": ["Not Started", "In Progress", "Complete", "On Hold"],
      "filterable": true
    },
    {
      "internalName": "DueDate",
      "displayName": "Due Date",
      "type": "DateTime",
      "filterable": true,
      "sortable": true
    },
    {
      "internalName": "Priority",
      "displayName": "Priority",
      "type": "Choice",
      "choices": ["High", "Medium", "Low"],
      "filterable": true
    },
    {
      "internalName": "Budget",
      "displayName": "Budget",
      "type": "Number",
      "filterable": true
    }
  ]
}`}
          </pre>
        </details>
      )}
    </div>
  );
};

export default ConfigSetup;
