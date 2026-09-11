import * as React from 'react';
import { useMemo, useState } from 'react';
import { TextField, Dropdown, IDropdownOption, Toggle, IconButton, Checkbox, ChoiceGroup, Text } from '@fluentui/react';
import { IRhinoConfig, IGridColumn, CellRenderer, GridDensity, RowClickAction, IFieldDefinition } from '../../../../shared/types/RhinoConfig';
import { fieldByName, systemFields, buildDefaultColumn, isSortableType, isFilterableType } from '../../../../shared/config/helpers';
import { sampleItems } from '../../../../shared/config/sample';
import DataGrid from '../../../../shared/components/DataGrid/DataGrid';
import FilterBar from '../../../../shared/components/FilterBar/FilterBar';
import { withConfig, moveItem } from '../DesignerApp';
import styles from '../DesignerApp.module.scss';

export interface IGridStepProps {
  config: IRhinoConfig;
  onChange: (next: IRhinoConfig) => void;
}

const RENDERERS: IDropdownOption[] = [
  { key: 'auto', text: 'Auto (by type)' },
  { key: 'text', text: 'Plain text' },
  { key: 'badge', text: 'Badge' },
  { key: 'tags', text: 'Tags' },
  { key: 'persona', text: 'Person avatar' },
  { key: 'date', text: 'Date' },
  { key: 'datetime', text: 'Date + time' },
  { key: 'number', text: 'Number' },
  { key: 'currency', text: 'Currency' },
  { key: 'percent', text: 'Percent' },
  { key: 'link', text: 'Link' },
  { key: 'check', text: 'Tick / cross' }
];

export const GridStep: React.FC<IGridStepProps> = ({ config, onChange }) => {
  const [previewSort, setPreviewSort] = useState<{ field: string; asc: boolean } | undefined>();
  const items = useMemo(() => sampleItems(config, 6), [config]);
  const grid = config.grid;

  const allFields: IFieldDefinition[] = config.fields.filter(f => !f.isSystem).concat(systemFields());
  const visible: Record<string, number> = {};
  grid.columns.forEach((c, i) => { visible[c.field] = i; });

  const setGrid = (mutate: (g: IRhinoConfig['grid']) => void): void => onChange(withConfig(config, d => mutate(d.grid)));
  const updateColumn = (ci: number, mutate: (c: IGridColumn) => void): void => setGrid(g => mutate(g.columns[ci]));

  const toggleColumn = (field: IFieldDefinition, on: boolean): void => setGrid(g => {
    if (on) { if (g.columns.every(c => c.field !== field.internalName)) g.columns.push(buildDefaultColumn(field)); }
    else g.columns = g.columns.filter(c => c.field !== field.internalName);
  });

  const sortOptions: IDropdownOption[] = [{ key: 'ID', text: 'ID (created order)' }].concat(
    allFields.filter(f => f.internalName !== 'ID' && isSortableType(f.type)).map(f => ({ key: f.internalName, text: f.displayName }))
  );

  // Preview-only sort of the sample rows so the header arrows feel real.
  const previewItems = useMemo(() => {
    if (!previewSort) return items;
    const f = fieldByName(config, previewSort.field);
    const copy = items.slice();
    copy.sort((a, b) => {
      const av = a[previewSort.field], bv = b[previewSort.field];
      const as = f?.type === 'User' ? String((av as { Title?: string })?.Title || '') : String(av ?? '');
      const bs = f?.type === 'User' ? String((bv as { Title?: string })?.Title || '') : String(bv ?? '');
      const r = typeof av === 'number' && typeof bv === 'number' ? av - bv : as.localeCompare(bs);
      return previewSort.asc ? r : -r;
    });
    return copy;
  }, [items, previewSort, config]);

  return (
    <div>
      <div className={styles.stepHeader}>
        <h2>Configure the grid</h2>
        <p>Choose which columns show, their order, sorting and filtering, and how the list behaves. The preview uses sample data.</p>
      </div>

      <div className={styles.split}>
        <div>
          <div className={styles.panel} style={{ marginBottom: 12 }}>
            <div className={styles.panelTitle}>Columns</div>
            <div className={styles.columnRow} style={{ borderBottom: '1px solid #e6e6e6' }}>
              <span /><span className={styles.colHead}>Field</span><span className={styles.colHead}>Label</span><span className={styles.colHead}>Width</span>
              <span className={styles.colHead}>Render</span><span className={styles.colHead}>Sort</span><span className={styles.colHead}>Filter</span><span />
            </div>
            {grid.columns.map((c, ci) => {
              const f = fieldByName(config, c.field);
              if (!f) return undefined;
              return (
                <div key={c.field} className={styles.columnRow}>
                  <Checkbox checked onChange={() => toggleColumn(f, false)} />
                  <div><div className={styles.placementName}>{f.displayName}{f.isSystem && <span className={styles.sysTag}>system</span>}</div><div className={styles.placementType}>{f.type}</div></div>
                  <TextField value={c.label || ''} placeholder={f.displayName} onChange={(_, v) => updateColumn(ci, x => { x.label = v || undefined; })} />
                  <TextField type="number" value={c.width ? String(c.width) : ''} placeholder="auto" onChange={(_, v) => updateColumn(ci, x => { x.width = v ? Number(v) : undefined; })} />
                  <Dropdown selectedKey={c.render || 'auto'} options={RENDERERS} onChange={(_, o) => updateColumn(ci, x => { x.render = o?.key as CellRenderer; })} />
                  <Toggle checked={c.sortable !== false && isSortableType(f.type)} disabled={!isSortableType(f.type)} onChange={(_, on) => updateColumn(ci, x => { x.sortable = !!on; })} styles={{ root: { marginBottom: 0 } }} />
                  <Toggle checked={!!c.filterable && isFilterableType(f.type)} disabled={!isFilterableType(f.type)} onChange={(_, on) => updateColumn(ci, x => { x.filterable = !!on; })} styles={{ root: { marginBottom: 0 } }} />
                  <div className={styles.rowBtns}>
                    <IconButton iconProps={{ iconName: 'Up' }} className={styles.iconBtn} disabled={ci === 0} onClick={() => setGrid(g => { g.columns = moveItem(g.columns, ci, ci - 1); })} />
                    <IconButton iconProps={{ iconName: 'Down' }} className={styles.iconBtn} disabled={ci === grid.columns.length - 1} onClick={() => setGrid(g => { g.columns = moveItem(g.columns, ci, ci + 1); })} />
                    <IconButton iconProps={{ iconName: c.align === 'right' ? 'AlignRight' : c.align === 'center' ? 'AlignCenter' : 'AlignLeft' }} className={styles.iconBtn} title="Alignment"
                      menuProps={{ items: [
                        { key: 'left', text: 'Left', onClick: () => updateColumn(ci, x => { x.align = undefined; }) },
                        { key: 'center', text: 'Centre', onClick: () => updateColumn(ci, x => { x.align = 'center'; }) },
                        { key: 'right', text: 'Right', onClick: () => updateColumn(ci, x => { x.align = 'right'; }) }
                      ] }} />
                  </div>
                </div>
              );
            })}
            {allFields.filter(f => visible[f.internalName] === undefined).length > 0 && (
              <div style={{ marginTop: 10 }}>
                <Text variant="small" className={styles.muted}>Hidden columns – tick to show:</Text>
                <div className={styles.chips} style={{ marginTop: 6 }}>
                  {allFields.filter(f => visible[f.internalName] === undefined).map(f => (
                    <button key={f.internalName} type="button" className={styles.chip} onClick={() => toggleColumn(f, true)}>＋ {f.displayName}</button>
                  ))}
                </div>
              </div>
            )}
          </div>

          <div className={styles.panel}>
            <div className={styles.panelTitle}>Behaviour</div>
            <div className={styles.settingsGrid}>
              <Dropdown label="Default sort" selectedKey={grid.defaultSort?.field || 'ID'} options={sortOptions}
                onChange={(_, o) => setGrid(g => { g.defaultSort = { field: String(o?.key), direction: g.defaultSort?.direction || 'asc' }; })} />
              <ChoiceGroup label="Direction" selectedKey={grid.defaultSort?.direction || 'desc'} styles={{ flexContainer: { display: 'flex', gap: 12 } }}
                options={[{ key: 'asc', text: 'Ascending' }, { key: 'desc', text: 'Descending' }]}
                onChange={(_, o) => setGrid(g => { g.defaultSort = { field: g.defaultSort?.field || 'ID', direction: o?.key as 'asc' | 'desc' }; })} />
              <Dropdown label="Page size" selectedKey={grid.pageSize || 50} options={[25, 50, 100, 200].map(n => ({ key: n, text: `${n} rows` }))}
                onChange={(_, o) => setGrid(g => { g.pageSize = Number(o?.key); })} />
              <Dropdown label="Density" selectedKey={grid.density || 'comfortable'} options={[{ key: 'comfortable', text: 'Comfortable' }, { key: 'compact', text: 'Compact' }]}
                onChange={(_, o) => setGrid(g => { g.density = o?.key as GridDensity; })} />
              <Dropdown label="Row click" selectedKey={grid.rowClick || 'edit'} options={[{ key: 'edit', text: 'Opens edit form' }, { key: 'view', text: 'Opens read-only view' }, { key: 'none', text: 'Does nothing' }]}
                onChange={(_, o) => setGrid(g => { g.rowClick = o?.key as RowClickAction; })} />
              <TextField label="Empty message" value={grid.emptyMessage || ''} placeholder="No items yet" onChange={(_, v) => setGrid(g => { g.emptyMessage = v || undefined; })} />
            </div>
            <div className={styles.togglesGrid} style={{ marginTop: 10 }}>
              <Toggle label="Search box" inlineLabel checked={grid.allowSearch !== false} onChange={(_, on) => setGrid(g => { g.allowSearch = !!on; })} />
              <Toggle label="Filter bar" inlineLabel checked={grid.allowFilter !== false} onChange={(_, on) => setGrid(g => { g.allowFilter = !!on; })} />
              <Toggle label="Column sorting" inlineLabel checked={grid.allowSort !== false} onChange={(_, on) => setGrid(g => { g.allowSort = !!on; })} />
              <Toggle label="Column resize" inlineLabel checked={grid.allowColumnResize !== false} onChange={(_, on) => setGrid(g => { g.allowColumnResize = !!on; })} />
              <Toggle label="Export to CSV" inlineLabel checked={grid.allowExport !== false} onChange={(_, on) => setGrid(g => { g.allowExport = !!on; })} />
              <Toggle label="Item count" inlineLabel checked={grid.showCount !== false} onChange={(_, on) => setGrid(g => { g.showCount = !!on; })} />
              <Toggle label="Allow add" inlineLabel checked={grid.allowAdd !== false} onChange={(_, on) => setGrid(g => { g.allowAdd = !!on; })} />
              <Toggle label="Allow edit" inlineLabel checked={grid.allowEdit !== false} onChange={(_, on) => setGrid(g => { g.allowEdit = !!on; })} />
              <Toggle label="Allow delete" inlineLabel checked={grid.allowDelete !== false} onChange={(_, on) => setGrid(g => { g.allowDelete = !!on; })} />
            </div>
          </div>
        </div>

        <div className={styles.previewFrame}>
          <div className={styles.previewLabel}><span>Live preview (sample data)</span></div>
          <div style={{ background: '#fff', borderRadius: 10, overflow: 'hidden', border: '1px solid #e6e6e6' }}>
            <DataGrid
              config={config}
              items={previewItems}
              loading={false}
              sortField={previewSort?.field || grid.defaultSort?.field || 'ID'}
              sortAsc={previewSort ? previewSort.asc : (grid.defaultSort?.direction || 'desc') === 'asc'}
              onSort={(f, asc) => setPreviewSort({ field: f, asc })}
              onAdd={() => window.alert('Preview only.')}
              onView={() => undefined}
              onEdit={() => undefined}
              onDelete={() => undefined}
              onRefresh={() => setPreviewSort(undefined)}
              filterBar={<FilterBar fields={config.fields} columns={grid.columns} filterState={{}} onFilterChange={() => undefined} />}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default GridStep;
