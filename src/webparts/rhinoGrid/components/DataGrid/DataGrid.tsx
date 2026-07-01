import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Selection,
  CommandBar,
  ICommandBarItemProps,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  Link,
  DefaultButton,
  Text
} from '@fluentui/react';
import { IFieldDefinition, IGridItem } from '../../types/IFieldConfig';
import styles from './DataGrid.module.scss';

export interface IDataGridProps {
  items: IGridItem[];
  fields: IFieldDefinition[];
  loading: boolean;
  error?: string;
  selectedItem?: IGridItem;
  onSelectionChange: (item: IGridItem | undefined) => void;
  onSort: (field: string, ascending: boolean) => void;
  onAdd: () => void;
  onEdit: () => void;
  onDelete: () => void;
  onRefresh: () => void;
  hasMore?: boolean;
  onLoadMore?: () => void;
}

export const DataGrid: React.FC<IDataGridProps> = ({
  items,
  fields,
  loading,
  error,
  selectedItem,
  onSelectionChange,
  onSort,
  onAdd,
  onEdit,
  onDelete,
  onRefresh,
  hasMore,
  onLoadMore
}) => {
  const [sortField, setSortField] = React.useState<string>('ID');
  const [sortAsc, setSortAsc] = React.useState<boolean>(true);

  const selection = React.useMemo(() => new Selection({
    onSelectionChanged: () => {
      const selected = selection.getSelection() as IGridItem[];
      onSelectionChange(selected.length > 0 ? selected[0] : undefined);
    }
  }), []);

  const handleColumnHeaderClick = (_: React.MouseEvent, column?: IColumn): void => {
    if (!column?.fieldName) return;
    const newAsc = column.fieldName === sortField ? !sortAsc : true;
    setSortField(column.fieldName);
    setSortAsc(newAsc);
    onSort(column.fieldName, newAsc);
  };

  const formatCellValue = (item: IGridItem, field: IFieldDefinition): JSX.Element | string => {
    const value = item[field.internalName];
    if (value === null || value === undefined) return '';

    switch (field.type) {
      case 'DateTime': {
        const date = new Date(value as string);
        return isNaN(date.getTime()) ? String(value) : date.toLocaleDateString();
      }
      case 'Boolean':
        return value ? (
          <Icon iconName="CheckMark" className={styles.boolTrue} />
        ) : (
          <Icon iconName="Cancel" className={styles.boolFalse} />
        );
      case 'User': {
        const userVal = value as { Title?: string };
        return userVal?.Title || '';
      }
      case 'URL': {
        const urlVal = value as { Url?: string; Description?: string };
        if (urlVal?.Url) {
          return <Link href={urlVal.Url} target="_blank">{urlVal.Description || urlVal.Url}</Link>;
        }
        return '';
      }
      case 'Number':
        return typeof value === 'number' ? value.toLocaleString() : String(value);
      default:
        return String(value);
    }
  };

  const columns: IColumn[] = fields.map(field => ({
    key: field.internalName,
    name: field.displayName,
    fieldName: field.internalName,
    minWidth: 80,
    maxWidth: 300,
    isResizable: true,
    isSorted: sortField === field.internalName,
    isSortedDescending: sortField === field.internalName && !sortAsc,
    onColumnClick: field.sortable ? handleColumnHeaderClick : undefined,
    onRender: (item: IGridItem) => formatCellValue(item, field)
  }));

  const commandItems: ICommandBarItemProps[] = [
    {
      key: 'add',
      text: 'Add',
      iconProps: { iconName: 'Add' },
      onClick: onAdd
    },
    {
      key: 'edit',
      text: 'Edit',
      iconProps: { iconName: 'Edit' },
      disabled: !selectedItem,
      onClick: onEdit
    },
    {
      key: 'delete',
      text: 'Delete',
      iconProps: { iconName: 'Delete' },
      disabled: !selectedItem,
      onClick: onDelete
    }
  ];

  const farCommandItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: onRefresh
    }
  ];

  return (
    <div className={styles.gridContainer}>
      <div className={styles.toolbar}>
        <CommandBar items={commandItems} farItems={farCommandItems} />
      </div>

      {error && (
        <div className={styles.errorState}>
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      )}

      <div className={styles.listContainer}>
        {items.length === 0 && !loading ? (
          <div className={styles.emptyState}>
            <Text variant="large">No items found.</Text>
            <br />
            <Text variant="medium">Use the Add button to create the first item, or adjust your filters.</Text>
          </div>
        ) : (
          <DetailsList
            items={items}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.single}
            selection={selection}
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
          <DefaultButton text="Load More" onClick={onLoadMore} disabled={loading} />
        </div>
      )}
    </div>
  );
};

export default DataGrid;
