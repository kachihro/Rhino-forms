import * as React from 'react';
import { useState, useMemo } from 'react';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, Spinner, SpinnerSize, MessageBar, MessageBarType, Icon,
  DefaultButton, PrimaryButton, IconButton, SearchBox, ContextualMenu, IContextualMenuItem, Text, ActionButton, IDetailsRowProps,
  DetailsRow, ConstrainMode
} from '@fluentui/react';
import { IRhinoConfig, IGridItem, IFieldDefinition, IGridColumn } from '../../types/RhinoConfig';
import { fieldByName, itemLabels, defaultColumnWidth } from '../../config/helpers';
import { renderCell } from '../cells/CellRenderers';
import { toPlainText } from '../cells/format';
import styles from './DataGrid.module.scss';

export interface IDataGridProps {
  config: IRhinoConfig;
  items: IGridItem[];
  loading: boolean;
  error?: string;
  sortField: string;
  sortAsc: boolean;
  onSort: (field: string, ascending: boolean) => void;
  onAdd: () => void;
  onView: (item: IGridItem) => void;
  onEdit: (item: IGridItem) => void;
  onDelete: (item: IGridItem) => void;
  onRefresh: () => void;
  hasMore?: boolean;
  onLoadMore?: () => void;
  filterBar?: React.ReactNode;
  /** Total loaded so far vs. filtered by search (footer text). */
  activeFilterCount?: number;
}

interface IResolvedColumn {
  column: IGridColumn;
  field: IFieldDefinition;
}

export const DataGrid: React.FC<IDataGridProps> = ({
  config, items, loading, error, sortField, sortAsc, onSort, onAdd, onView, onEdit, onDelete, onRefresh, hasMore, onLoadMore, filterBar
}) => {
  const grid = config.grid;
  const [searchText, setSearchText] = useState('');
  const [menu, setMenu] = useState<{ target: HTMLElement; item: IGridItem } | undefined>();
  const { singular, plural } = itemLabels(config);

  const resolved: IResolvedColumn[] = useMemo(() =>
    grid.columns
      .map(c => ({ column: c, field: fieldByName(config, c.field) as IFieldDefinition }))
      .filter(x => !!x.field),
    [grid.columns, config.fields]);

  const filteredItems = useMemo(() => {
    const term = searchText.trim().toLowerCase();
    if (!term) return items;
    return items.filter(item => resolved.some(({ field }) => toPlainText(item[field.internalName], field).toLowerCase().indexOf(term) >= 0));
  }, [items, searchText, resolved]);

  const handleColumnClick = (_: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
    if (!column?.fieldName) return;
    const asc = column.fieldName === sortField ? !sortAsc : true;
    onSort(column.fieldName, asc);
  };

  const exportCsv = (): void => {
    const esc = (s: string): string => `"${s.replace(/"/g, '""')}"`;
    const header = resolved.map(r => esc(r.column.label || r.field.displayName)).join(',');
    const rows = filteredItems.map(item => resolved.map(r => esc(toPlainText(item[r.field.internalName], r.field))).join(','));
    const csv = '﻿' + [header].concat(rows).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${(config.branding.title || config.list.title || plural).replace(/[^\w\- ]+/g, '')}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const columns: IColumn[] = resolved.map(({ column, field }) => {
    const sortable = grid.allowSort !== false && column.sortable !== false && field.type !== 'MultiChoice' && field.type !== 'URL';
    const min = column.width || defaultColumnWidth(field.type);
    return {
      key: field.internalName,
      name: column.label || field.displayName,
      fieldName: field.internalName,
      minWidth: min,
      maxWidth: column.maxWidth || Math.max(min * 2, field.type === 'Note' ? 420 : 260),
      isResizable: grid.allowColumnResize !== false,
      isSorted: sortField === field.internalName,
      isSortedDescending: sortField === field.internalName && !sortAsc,
      onColumnClick: sortable ? handleColumnClick : undefined,
      headerClassName: sortable ? styles.sortableHeader : styles.staticHeader,
      className: column.align === 'right' ? styles.alignRight : column.align === 'center' ? styles.alignCenter : undefined,
      onRender: (item: IGridItem) => renderCell(item, field, column)
    };
  });

  const showActions = grid.allowEdit !== false || grid.allowDelete !== false || grid.rowClick !== 'none';
  if (showActions) {
    columns.push({
      key: '__actions',
      name: '',
      fieldName: '__actions',
      minWidth: 70,
      maxWidth: 70,
      isResizable: false,
      onRender: (item: IGridItem) => (
        <div className={styles.rowActions}>
          {grid.allowEdit !== false && (
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" className={styles.rowActionBtn}
              onClick={e => { e.stopPropagation(); onEdit(item); }} />
          )}
          <IconButton iconProps={{ iconName: 'MoreVertical' }} title="More" className={styles.rowActionBtn}
            onClick={e => { e.stopPropagation(); setMenu({ target: e.currentTarget as HTMLElement, item }); }} />
        </div>
      )
    });
  }

  const onRowClick = (item: IGridItem): void => {
    if (grid.rowClick === 'edit' && grid.allowEdit !== false) onEdit(item);
    else if (grid.rowClick === 'view' || grid.rowClick === 'edit') onView(item);
  };

  const renderRow = (props?: IDetailsRowProps): JSX.Element | null => {
    if (!props) return null;
    const clickable = grid.rowClick !== 'none';
    return (
      <div className={clickable ? styles.clickableRow : undefined} onClick={clickable ? () => onRowClick(props.item as IGridItem) : undefined}>
        <DetailsRow {...props} className={grid.density === 'compact' ? styles.compactRow : styles.comfortableRow} />
      </div>
    );
  };

  const countText = searchText
    ? `${filteredItems.length} of ${items.length} ${plural}`
    : `${items.length}${hasMore ? '+' : ''} ${items.length === 1 ? singular : plural}`;

  const menuItems: IContextualMenuItem[] = menu ? [
    { key: 'view', text: 'View', iconProps: { iconName: 'View' }, onClick: () => { onView(menu.item); setMenu(undefined); } },
    ...(grid.allowEdit !== false ? [{ key: 'edit', text: 'Edit', iconProps: { iconName: 'Edit' }, onClick: () => { onEdit(menu.item); setMenu(undefined); } }] : []),
    ...(grid.allowDelete !== false ? [{ key: 'delete', text: 'Delete', iconProps: { iconName: 'Delete' }, className: styles.deleteMenuItem, onClick: () => { onDelete(menu.item); setMenu(undefined); } }] : [])
  ] : [];

  return (
    <div className={styles.gridContainer}>
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          {grid.allowAdd !== false && (
            <PrimaryButton text={`New ${singular}`} iconProps={{ iconName: 'Add' }} onClick={onAdd} className={styles.newBtn} />
          )}
          <ActionButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={onRefresh} />
          {grid.allowExport !== false && <ActionButton iconProps={{ iconName: 'ExcelDocument' }} text="Export" onClick={exportCsv} disabled={filteredItems.length === 0} />}
        </div>
        <div className={styles.toolbarRight}>
          {grid.allowSearch !== false && (
            <SearchBox placeholder={`Search ${plural}`} value={searchText} onChange={(_, v) => setSearchText(v || '')}
              onClear={() => setSearchText('')} className={styles.searchBox} />
          )}
        </div>
      </div>

      {(grid.allowFilter !== false && filterBar) || grid.showCount !== false ? (
        <div className={styles.filterRow}>
          <div className={styles.filterSlot}>{grid.allowFilter !== false ? filterBar : undefined}</div>
          {grid.showCount !== false && <Text className={styles.countText}>{countText}</Text>}
        </div>
      ) : undefined}

      {error && <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar}>{error}</MessageBar>}

      <div className={styles.listContainer}>
        {filteredItems.length === 0 && !loading ? (
          <div className={styles.emptyState}>
            <Icon iconName={searchText ? 'SearchIssue' : 'Inbox'} className={styles.emptyIcon} />
            <Text variant="large" className={styles.emptyTitle}>{searchText ? `No ${plural} match "${searchText}"` : (grid.emptyMessage || `No ${plural} yet`)}</Text>
            <Text variant="medium" className={styles.emptySubtitle}>
              {searchText ? 'Try a different search or clear the filters.' : (grid.allowAdd !== false ? `Use "New ${singular}" to add the first one.` : '')}
            </Text>
          </div>
        ) : (
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            selectionMode={SelectionMode.none}
            compact={grid.density === 'compact'}
            setKey="Id"
            getKey={(item: IGridItem) => String(item.Id)}
            isHeaderVisible
            onRenderRow={renderRow}
            styles={{ root: { overflowX: 'auto' } }}
          />
        )}
        {loading && (
          <div className={styles.loadingOverlay}>
            <Spinner size={SpinnerSize.large} label="Loading…" />
          </div>
        )}
      </div>

      {(hasMore && onLoadMore) && (
        <div className={styles.loadMore}>
          <DefaultButton text={`Load ${grid.pageSize || 50} more`} iconProps={{ iconName: 'ChevronDown' }} onClick={onLoadMore} disabled={loading} />
        </div>
      )}

      {menu && (
        <ContextualMenu target={menu.target} onDismiss={() => setMenu(undefined)} items={menuItems} />
      )}
    </div>
  );
};

export default DataGrid;
