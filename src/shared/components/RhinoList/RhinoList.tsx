import * as React from 'react';
import { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Icon, Text } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { IList } from '@pnp/sp/lists';
import { IRhinoConfig, IFilterState, IGridItem } from '../../types/RhinoConfig';
import { itemLabels } from '../../config/helpers';
import { GridDataService } from '../../services/GridDataService';
import { ListSetupService } from '../../services/ListSetupService';
import { PeopleService } from '../../services/PeopleService';
import DataGrid from '../DataGrid/DataGrid';
import FilterBar from '../FilterBar/FilterBar';
import ItemForm, { FormMode } from '../ItemForm/ItemForm';
import DeleteConfirmDialog from '../DeleteConfirmDialog/DeleteConfirmDialog';
import styles from './RhinoList.module.scss';

export interface IRhinoListProps {
  sp: SPFI;
  config: IRhinoConfig;
  /** Shown top-right of the header (e.g. an "Edit config" link in page edit mode). */
  headerExtra?: React.ReactNode;
  /** Hide the branded header card. */
  hideHeader?: boolean;
}

function hexToSoft(hex?: string): string | undefined {
  if (!hex || !/^#[0-9a-f]{6}$/i.test(hex)) return undefined;
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, 0.10)`;
}

/**
 * The runtime list experience: header + grid + add/edit/view form + delete,
 * fully driven by an IRhinoConfig.
 */
export const RhinoList: React.FC<IRhinoListProps> = ({ sp, config, headerExtra, hideHeader }) => {
  const [list, setList] = useState<IList | undefined>();
  const [resolveError, setResolveError] = useState<string | undefined>();
  const [items, setItems] = useState<IGridItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [filterState, setFilterState] = useState<IFilterState>({});
  const [sortField, setSortField] = useState<string>(config.grid.defaultSort?.field || 'ID');
  const [sortAsc, setSortAsc] = useState<boolean>((config.grid.defaultSort?.direction || 'desc') === 'asc');
  const [hasMore, setHasMore] = useState(false);
  const [formMode, setFormMode] = useState<FormMode | undefined>();
  const [selected, setSelected] = useState<IGridItem | undefined>();
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [toast, setToast] = useState<string | undefined>();
  const requestId = useRef(0);

  const pageSize = config.grid.pageSize || 50;
  const setup = useMemo(() => new ListSetupService(sp), [sp]);
  const people = useMemo(() => new PeopleService(sp), [sp]);
  const data = useMemo(() => list ? new GridDataService(list) : undefined, [list]);
  const { singular } = itemLabels(config);

  // Resolve the list once per config.
  useEffect(() => {
    let cancelled = false;
    setList(undefined);
    setResolveError(undefined);
    setup.resolveList(config.list)
      .then(l => { if (!cancelled) setList(l); })
      .catch(() => { if (!cancelled) setResolveError(`The list "${config.list.title}" could not be found on this site. Open the Rhino List Designer to provision it.`); });
    return () => { cancelled = true; };
  }, [setup, config.list.title, config.list.internalName]);

  useEffect(() => {
    setSortField(config.grid.defaultSort?.field || 'ID');
    setSortAsc((config.grid.defaultSort?.direction || 'desc') === 'asc');
    setFilterState({});
  }, [config.grid.defaultSort?.field, config.grid.defaultSort?.direction]);

  const load = useCallback(async (skip: number, append: boolean) => {
    if (!data) return;
    const id = ++requestId.current;
    setLoading(true);
    setError(undefined);
    try {
      const page = await data.getPage(config, { filters: filterState, sortField, sortAsc, top: pageSize, skip });
      if (id !== requestId.current) return;
      setItems(prev => append ? prev.concat(page.items) : page.items);
      setHasMore(page.hasMore);
    } catch (e) {
      if (id !== requestId.current) return;
      setError(`Could not load ${itemLabels(config).plural}: ${(e as Error).message}`);
    } finally {
      if (id === requestId.current) setLoading(false);
    }
  }, [data, config, filterState, sortField, sortAsc, pageSize]);

  useEffect(() => { void load(0, false); }, [load]);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(undefined), 3500);
    return () => clearTimeout(t);
  }, [toast]);

  const refresh = (): void => { void load(0, false); };

  const handleSave = async (payload: Record<string, unknown>): Promise<void> => {
    if (!data) return;
    if (formMode === 'add') {
      await data.add(payload);
      setToast(`${singular.charAt(0).toUpperCase()}${singular.slice(1)} added`);
    } else if (selected) {
      await data.update(selected.Id, payload);
      setToast('Changes saved');
    }
    setFormMode(undefined);
    setSelected(undefined);
    refresh();
  };

  const handleDelete = async (): Promise<void> => {
    if (!data || !selected) return;
    setDeleting(true);
    try {
      await data.delete(selected.Id);
      setConfirmDelete(false);
      setFormMode(undefined);
      setSelected(undefined);
      setToast(`${singular.charAt(0).toUpperCase()}${singular.slice(1)} deleted`);
      refresh();
    } catch (e) {
      setError(`Could not delete: ${(e as Error).message}`);
      setConfirmDelete(false);
    } finally {
      setDeleting(false);
    }
  };

  const openForm = async (mode: FormMode, item?: IGridItem): Promise<void> => {
    setSelected(item);
    setFormMode(mode);
    // Refresh the single item so the form shows the latest server state.
    if (item && data) {
      try {
        const fresh = await data.getById(config, item.Id);
        setSelected(fresh);
      } catch { /* keep the grid copy */ }
    }
  };

  const accent = config.branding.accent;
  const cssVars: React.CSSProperties = accent ? ({ '--rhinoAccent': accent, '--rhinoAccentSoft': hexToSoft(accent) } as React.CSSProperties) : {};
  const iconValue = config.branding.icon || '🦏';
  const isFluentIcon = /^[A-Za-z][A-Za-z0-9]+$/.test(iconValue) && iconValue.length > 2;

  return (
    <div className={styles.app} style={cssVars}>
      <div className={styles.card}>
        {!hideHeader && (
          <div className={styles.header}>
            <div className={styles.headerIcon}>{isFluentIcon ? <Icon iconName={iconValue} /> : iconValue}</div>
            <div className={styles.headerText}>
              <div className={styles.headerTitle}>{config.branding.title || config.list.title}</div>
              {config.branding.subtitle && <div className={styles.headerSubtitle}>{config.branding.subtitle}</div>}
            </div>
            <div className={styles.headerRight}>
              {toast && <span className={styles.toast}><Icon iconName="CheckMark" /> {toast}</span>}
              {headerExtra}
            </div>
          </div>
        )}

        {resolveError && (
          <MessageBar messageBarType={MessageBarType.warning} className={styles.resolveError}>{resolveError}</MessageBar>
        )}
        {!list && !resolveError && (
          <div className={styles.resolving}><Spinner size={SpinnerSize.medium} label="Connecting to list…" /></div>
        )}

        {list && (
          <DataGrid
            config={config}
            items={items}
            loading={loading}
            error={error}
            sortField={sortField}
            sortAsc={sortAsc}
            onSort={(f, asc) => { setSortField(f); setSortAsc(asc); }}
            onAdd={() => { void openForm('add'); }}
            onView={item => { void openForm('view', item); }}
            onEdit={item => { void openForm('edit', item); }}
            onDelete={item => { setSelected(item); setConfirmDelete(true); }}
            onRefresh={refresh}
            hasMore={hasMore}
            onLoadMore={() => { void load(items.length, true); }}
            filterBar={
              <FilterBar fields={config.fields} columns={config.grid.columns} filterState={filterState} onFilterChange={setFilterState} />
            }
          />
        )}
      </div>

      {formMode && (
        <ItemForm
          config={config}
          mode={formMode}
          item={formMode === 'add' ? undefined : selected}
          peopleService={people}
          onSave={handleSave}
          onDismiss={() => { setFormMode(undefined); setSelected(undefined); }}
          onEdit={() => setFormMode('edit')}
          onDelete={() => setConfirmDelete(true)}
        />
      )}

      {confirmDelete && selected && (
        <DeleteConfirmDialog
          itemTitle={String(selected.Title || `${singular} #${selected.Id}`)}
          itemLabel={singular}
          busy={deleting}
          onConfirm={() => { void handleDelete(); }}
          onDismiss={() => setConfirmDelete(false)}
        />
      )}

      {!hideHeader && (
        <div className={styles.footerNote}><Text variant="tiny">Rhino Forms · {config.list.title}</Text></div>
      )}
    </div>
  );
};

export default RhinoList;
