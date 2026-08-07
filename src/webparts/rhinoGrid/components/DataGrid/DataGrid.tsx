import * as React from 'react';
import { useState, useMemo } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  Link,
  DefaultButton,
  PrimaryButton,
  IconButton,
  SearchBox,
  ContextualMenu,
  IContextualMenuItem,
  Text,
  ActionButton
} from '@fluentui/react';
import { IFieldDefinition, IGridItem } from '../../types/IFieldConfig';
import styles from './DataGrid.module.scss';

function getInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : name.slice(0, 2).toUpperCase();
}

const AVATAR_COLORS = ['#0078d4', '#107c10', '#881798', '#ca5010', '#038387', '#8764b8', '#004b50'];

function getAvatarColor(name: string): string {
  let h = 5381;
  for (let i = 0; i < name.length; i++) h = ((h << 5) + h) ^ name.charCodeAt(i);
  return AVATAR_COLORS[Math.abs(h) % AVATAR_COLORS.length];
}

function getBadgeColors(value: string): { bg: string; color: string; dot: string } {
  const v = value.toLowerCase();
  if (/pend|progress|wait|review|submitted/.test(v)) return { bg: '#fef7e0', color: '#6a4204', dot: '#c19c00' };
  if (/approv|complet|active|done|success|paid/.test(v)) return { bg: '#dff6dd', color: '#107c10', dot: '#107c10' };
  if (/reject|cancel|fail|deni|close/.test(v)) return { bg: '#fde7e9', color: '#a4262c', dot: '#a4262c' };
  if (/hold|paused|defer|suspend/.test(v)) return { bg: '#e8e8e8', color: '#323130', dot: '#605e5c' };
  return { bg: '#deecf9', color: '#004578', dot: '#0078d4' };
}

export interface IDataGridProps {
  items: IGridItem[];
  fields: IFieldDefinition[];
  loading: boolean;
  error?: string;
  onSort: (field: string, ascending: boolean) => void;
  onAdd: () => void;
  onEditItem: (item: IGridItem) => void;
  onDeleteItem: (item: IGridItem) => void;
  onRefresh: () => void;
  hasMore?: boolean;
  onLoadMore?: () => void;
  itemLabel?: string;
  filterBar?: React.ReactNode;
}

export const DataGrid: React.FC<IDataGridProps> = ({
  items,
  fields,
  loading,
  error,
  onSort,
  onAdd,
  onEditItem,
  onDeleteItem,
  onRefresh,
  hasMore,
  onLoadMore,
  itemLabel = 'item',
  filterBar
}) => {
  const [sortField, setSortField] = useState<string>('ID');
  const [sortAsc, setSortAsc] = useState<boolean>(true);
  const [searchText, setSearchText] = useState('');
  const [menuTarget, setMenuTarget] = useState<HTMLElement | null>(null);
  const [menuItem, setMenuItem] = useState<IGridItem | null>(null);

  const filteredItems = useMemo(() => {
    if (!searchText.trim()) return items;
    const term = searchText.toLowerCase();
    return items.filter(item =>
      fields.some(f => {
        const val = item[f.internalName];
        if (val === null || val === undefined) return false;
        if (f.type === 'User') return String((val as { Title?: string })?.Title || '').toLowerCase().includes(term);
        if (f.type === 'Choice' || f.type === 'Text' || f.type === 'Note') return String(val).toLowerCase().includes(term);
        return false;
      })
    );
  }, [items, searchText, fields]);

  const handleColumnClick = (_: React.MouseEvent, column?: IColumn): void => {
    if (!column?.fieldName) return;
    const asc = column.fieldName === sortField ? !sortAsc : true;
    setSortField(column.fieldName);
    setSortAsc(asc);
    onSort(column.fieldName, asc);
  };

  const renderCell = (item: IGridItem, field: IFieldDefinition): JSX.Element | string => {
    const value = item[field.internalName];
    if (value === null || value === undefined) return '';

    switch (field.type) {
      case 'DateTime': {
        const d = new Date(value as string);
        return isNaN(d.getTime()) ? String(value) : d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      }
      case 'Boolean':
        return value
          ? <Icon iconName="CheckMark" className={styles.boolTrue} />
          : <Icon iconName="Cancel" className={styles.boolFalse} />;
      case 'User': {
        const u = value as { Title?: string };
        const name = u?.Title || '';
        if (!name) return '';
        return (
          <div className={styles.persona}>
            <div className={styles.avatar} style={{ background: getAvatarColor(name) }}>{getInitials(name)}</div>
            <span className={styles.personaName}>{name}</span>
          </div>
        );
      }
      case 'URL': {
        const u = value as { Url?: string; Description?: string };
        return u?.Url ? <Link href={u.Url} target="_blank">{u.Description || u.Url}</Link> : '';
      }
      case 'Choice': {
        const str = String(value);
        const { bg, color, dot } = getBadgeColors(str);
        return (
          <span className={styles.badge} style={{ background: bg, color }}>
            <span className={styles.badgeDot} style={{ background: dot }} />
            {str}
          </span>
        );
      }
      case 'Number':
        return typeof value === 'number' ? value.toLocaleString() : String(value);
      default:
        return String(value);
    }
  };

  const exportCsv = (): void => {
    const header = fields.map(f => `"${f.displayName}"`).join(',');
    const rows = filteredItems.map(item =>
      fields.map(f => {
        const v = item[f.internalName];
        if (v === null || v === undefined) return '""';
        if (f.type === 'User') return `"${(v as { Title?: string })?.Title || ''}"`;
        if (f.type === 'DateTime') return `"${new Date(v as string).toLocaleDateString()}"`;
        if (f.type === 'URL') return `"${(v as { Url?: string })?.Url || ''}"`;
        return `"${String(v).replace(/"/g, '""')}"`;
      }).join(',')
    );
    const csv = [header, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${itemLabel}s.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const columns: IColumn[] = [
    ...fields.map(field => ({
      key: field.internalName,
      name: field.displayName,
      fieldName: field.internalName,
      minWidth: field.type === 'User' ? 160 : field.type === 'Note' ? 200 : 100,
      maxWidth: field.type === 'Note' ? 350 : 260,
      isResizable: true,
      isSorted: sortField === field.internalName,
      isSortedDescending: sortField === field.internalName && !sortAsc,
      onColumnClick: field.sortable ? handleColumnClick : undefined,
      onRender: (item: IGridItem) => renderCell(item, field)
    })),
    {
      key: '__actions',
      name: '',
      fieldName: '__actions',
      minWidth: 64,
      maxWidth: 64,
      isResizable: false,
      onRender: (item: IGridItem) => (
        <div className={styles.rowActions}>
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            title="Edit"
            className={styles.rowActionBtn}
            onClick={(e) => { e.stopPropagation(); onEditItem(item); }}
          />
          <IconButton
            iconProps={{ iconName: 'MoreVertical' }}
            title="More options"
            className={styles.rowActionBtn}
            onClick={(e) => { e.stopPropagation(); setMenuTarget(e.currentTarget as HTMLElement); setMenuItem(item); }}
          />
        </div>
      )
    }
  ];

  const pluralLabel = `${itemLabel}s`;
  const countText = searchText
    ? `${filteredItems.length} of ${items.length} ${pluralLabel}`
    : `${items.length} ${pluralLabel}`;

  return (
    <div className={styles.gridContainer}>
      {/* Toolbar */}
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          <PrimaryButton
            text={`+ New ${itemLabel}`}
            onClick={onAdd}
            className={styles.newBtn}
          />
          <ActionButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={onRefresh} />
          <ActionButton iconProps={{ iconName: 'Download' }} text="Export" onClick={exportCsv} />
        </div>
        <div className={styles.toolbarRight}>
          <SearchBox
            placeholder={`Search ${pluralLabel}`}
            value={searchText}
            onChange={(_, v) => setSearchText(v || '')}
            onClear={() => setSearchText('')}
            className={styles.searchBox}
          />
        </div>
      </div>

      {/* Filter bar slot + count */}
      <div className={styles.filterRow}>
        <div className={styles.filterSlot}>{filterBar}</div>
        <div className={styles.countBadge}>
          <Text className={styles.countText}>{countText}</Text>
        </div>
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar}>{error}</MessageBar>
      )}

      {/* Grid */}
      <div className={styles.listContainer}>
        {filteredItems.length === 0 && !loading ? (
          <div className={styles.emptyState}>
            <Icon iconName="SearchIssue" className={styles.emptyIcon} />
            <Text variant="large" className={styles.emptyTitle}>No {pluralLabel} found</Text>
            <Text variant="medium" className={styles.emptySubtitle}>
              Try adjusting your search or filters, or add a new {itemLabel}.
            </Text>
          </div>
        ) : (
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            setKey="Id"
            isHeaderVisible
          />
        )}

        {loading && (
          <div className={styles.loadingOverlay}>
            <Spinner size={SpinnerSize.large} label="Loading..." />
          </div>
        )}
      </div>

      {hasMore && onLoadMore && (
        <div className={styles.loadMoreContainer}>
          <DefaultButton text="Load more" onClick={onLoadMore} disabled={loading} />
        </div>
      )}

      <div className={styles.footer}>
        <Text className={styles.footerText}>Showing {filteredItems.length} of {items.length} {pluralLabel}</Text>
      </div>

      {menuItem && menuTarget && (
        <ContextualMenu
          target={menuTarget}
          onDismiss={() => { setMenuTarget(null); setMenuItem(null); }}
          items={[
            {
              key: 'edit',
              text: 'Edit',
              iconProps: { iconName: 'Edit' },
              onClick: () => { onEditItem(menuItem); setMenuTarget(null); setMenuItem(null); }
            },
            {
              key: 'delete',
              text: 'Delete',
              iconProps: { iconName: 'Delete' },
              className: styles.deleteMenuItem,
              onClick: () => { onDeleteItem(menuItem); setMenuTarget(null); setMenuItem(null); }
            }
          ] as IContextualMenuItem[]}
        />
      )}
    </div>
  );
};

export default DataGrid;
