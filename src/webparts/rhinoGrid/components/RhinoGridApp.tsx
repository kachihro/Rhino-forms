import * as React from 'react';
import { useState, useEffect, useCallback, useMemo } from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IGridConfig, IFilterState, IGridItem } from '../types/IFieldConfig';
import { GridDataService } from '../services/GridDataService';
import { ListSetupService } from '../services/ListSetupService';
import ConfigSetup from './ConfigSetup/ConfigSetup';
import DataGrid from './DataGrid/DataGrid';
import FilterBar from './FilterBar/FilterBar';
import ItemForm from './ItemForm/ItemForm';
import DeleteConfirmDialog from './DeleteConfirmDialog/DeleteConfirmDialog';
import styles from './RhinoGridApp.module.scss';

export interface IRhinoGridAppProps {
  sp: SPFI;
  config: IGridConfig | undefined;
  displayMode: DisplayMode;
  onConfigSaved: (json: string, listName: string) => void;
}

const PAGE_SIZE = 100;

const RhinoGridApp: React.FC<IRhinoGridAppProps> = ({ sp, config, displayMode, onConfigSaved }) => {
  const [items, setItems] = useState<IGridItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [filterState, setFilterState] = useState<IFilterState>({});
  const [selectedItem, setSelectedItem] = useState<IGridItem | undefined>();
  const [formMode, setFormMode] = useState<'add' | 'edit' | null>(null);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [listProvisioned, setListProvisioned] = useState(false);
  const [checkingList, setCheckingList] = useState(false);
  const [sortField, setSortField] = useState<string>('ID');
  const [sortAsc, setSortAsc] = useState(true);
  const [skip, setSkip] = useState(0);
  const [hasMore, setHasMore] = useState(false);

  const isReadMode = displayMode === DisplayMode.Read;

  const dataService = useMemo(() => config ? new GridDataService(sp) : null, [sp, config?.listName]);
  const setupService = useMemo(() => new ListSetupService(sp), [sp]);

  useEffect(() => {
    if (config) {
      setCheckingList(true);
      setupService.listExists(config.listName)
        .then(exists => setListProvisioned(exists))
        .catch(() => setListProvisioned(false))
        .finally(() => setCheckingList(false));
    } else {
      setListProvisioned(false);
    }
  }, [config?.listName]);

  const loadItems = useCallback(async (currentSkip: number = 0) => {
    if (!config || !dataService || !listProvisioned) return;
    setLoading(true);
    setError(undefined);
    try {
      const loaded = await dataService.getItems(
        config.listName,
        config.fields,
        filterState,
        sortField,
        sortAsc,
        PAGE_SIZE + 1,
        currentSkip
      );
      const hasMoreItems = loaded.length > PAGE_SIZE;
      const pageItems = hasMoreItems ? loaded.slice(0, PAGE_SIZE) : loaded;
      if (currentSkip === 0) {
        setItems(pageItems);
      } else {
        setItems(prev => [...prev, ...pageItems]);
      }
      setHasMore(hasMoreItems);
    } catch (e) {
      setError(`Failed to load items: ${e.message}`);
    } finally {
      setLoading(false);
    }
  }, [config, dataService, filterState, sortField, sortAsc, listProvisioned]);

  useEffect(() => {
    if (listProvisioned) {
      setSkip(0);
      void loadItems(0);
    }
  }, [filterState, sortField, sortAsc, listProvisioned]);

  const handleFilterChange = (newState: IFilterState): void => {
    setFilterState(newState);
    setSkip(0);
  };

  const handleSort = (field: string, ascending: boolean): void => {
    setSortField(field);
    setSortAsc(ascending);
    setSkip(0);
  };

  const handleLoadMore = (): void => {
    const newSkip = skip + PAGE_SIZE;
    setSkip(newSkip);
    void loadItems(newSkip);
  };

  const handleRefresh = (): void => {
    setSkip(0);
    void loadItems(0);
  };

  const handleConfigSaved = (json: string, listName: string): void => {
    setListProvisioned(true);
    onConfigSaved(json, listName);
  };

  const handleEditItem = (item: IGridItem): void => {
    setSelectedItem(item);
    setFormMode('edit');
  };

  const handleDeleteItem = (item: IGridItem): void => {
    setSelectedItem(item);
    setShowDeleteDialog(true);
  };

  const handleDeleteConfirm = async (): Promise<void> => {
    if (!selectedItem || !config || !dataService) return;
    try {
      await dataService.deleteItem(config.listName, selectedItem.Id);
      setShowDeleteDialog(false);
      setSelectedItem(undefined);
      setSkip(0);
      void loadItems(0);
    } catch (e) {
      setError(`Failed to delete item: ${e.message}`);
      setShowDeleteDialog(false);
    }
  };

  if (checkingList) {
    return (
      <div className={styles.appContainer}>
        <MessageBar>Checking SharePoint list status...</MessageBar>
      </div>
    );
  }

  if (!config) {
    if (isReadMode) {
      return (
        <div className={styles.appContainer}>
          <div className={styles.readOnlyMessage}>
            <MessageBar messageBarType={MessageBarType.info}>
              This webpart needs to be configured. Edit the page and upload a JSON configuration file.
            </MessageBar>
          </div>
        </div>
      );
    }
    return (
      <div className={styles.appContainer}>
        <ConfigSetup sp={sp} setupService={setupService} onConfigSaved={handleConfigSaved} />
      </div>
    );
  }

  if (!listProvisioned) {
    if (isReadMode) {
      return (
        <div className={styles.appContainer}>
          <div className={styles.readOnlyMessage}>
            <MessageBar messageBarType={MessageBarType.warning}>
              The SharePoint list &ldquo;{config.listName}&rdquo; has not been provisioned yet. Edit the page to provision it.
            </MessageBar>
          </div>
        </div>
      );
    }
    return (
      <div className={styles.appContainer}>
        <ConfigSetup sp={sp} setupService={setupService} existingConfig={config} onConfigSaved={handleConfigSaved} />
      </div>
    );
  }

  const itemLabel = config.itemLabel || 'item';

  return (
    <div className={styles.appContainer}>
      <div className={styles.card}>
        <div className={styles.header}>
          <div className={styles.headerIcon}>{config.icon || '🦏'}</div>
          <div className={styles.headerText}>
            <div className={styles.headerTitle}>{config.listName}</div>
            {config.listDescription && (
              <div className={styles.headerSubtitle}>{config.listDescription}</div>
            )}
          </div>
        </div>

        <DataGrid
          items={items}
          fields={config.fields}
          loading={loading}
          error={error}
          onSort={handleSort}
          onAdd={() => { setSelectedItem(undefined); setFormMode('add'); }}
          onEditItem={handleEditItem}
          onDeleteItem={handleDeleteItem}
          onRefresh={handleRefresh}
          hasMore={hasMore}
          onLoadMore={handleLoadMore}
          itemLabel={itemLabel}
          filterBar={
            <FilterBar
              fields={config.fields}
              filterState={filterState}
              onFilterChange={handleFilterChange}
            />
          }
        />
      </div>

      {formMode && (
        <ItemForm
          mode={formMode}
          fields={config.fields}
          item={formMode === 'edit' ? selectedItem : undefined}
          listName={config.listName}
          dataService={dataService!}
          onSaved={() => { setFormMode(null); handleRefresh(); }}
          onDismiss={() => setFormMode(null)}
        />
      )}

      {showDeleteDialog && selectedItem && (
        <DeleteConfirmDialog
          itemTitle={String(selectedItem.Title || `Item #${selectedItem.Id}`)}
          onConfirm={handleDeleteConfirm}
          onDismiss={() => setShowDeleteDialog(false)}
        />
      )}
    </div>
  );
};

export default RhinoGridApp;
