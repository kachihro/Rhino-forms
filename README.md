# Rhino Forms

A config-driven replacement for PowerApps list forms on SharePoint Online. Two SPFx web parts, one JSON file:

| Web part | What it does |
| --- | --- |
| **Rhino List Designer** | Pick or define a list (with a proper internal name), provision it, lay out the form, configure the grid, save the JSON config to Site Assets. |
| **Rhino List** | Reads that JSON config and renders a fast, filterable, sortable grid with a beautiful add / edit / view form. |

Everything the end user sees is described by a single `IRhinoConfig` JSON document. Edit it in the Designer, in the file, or in source control.

## How it fits together

1. Add **Rhino List Designer** to a page (a full-page app page works best).
2. **1. List** – choose *New list* or *Existing list*. Set the title and the **internal name** (this becomes `/Lists/<internalName>`). Click **Provision list**.
3. **2. Fields** – add columns. Display names are free text; internal names are generated (`Due Date` → `DueDate`, never `_x0020_`) and editable.
4. **3. Form layout** – sections, 1 to 3 columns, field widths, label / placeholder / help overrides, read-only, collapsible sections. Live preview.
5. **4. Grid** – visible columns and order, sortable / filterable per column, renderer (badge, persona, currency…), default sort, page size, density, row-click behaviour, feature toggles. Live preview.
6. **5. JSON** – view or hand-edit the config. **Save to site** writes `SiteAssets/RhinoForms/<internalName>.json`.
7. Add a **Rhino List** web part to any page and pick the config file in the property pane.

Re-running **Provision list** is safe: it adds missing columns and syncs titles, descriptions, required flags and choices. It never deletes anything.

## The config file

See [`samples/project-tracker.json`](samples/project-tracker.json) for a complete example. Minimal hand-written configs work too; anything omitted gets a sensible default, and the original v1 shape (`{ "listName": …, "fields": [...] }`) still loads.

```jsonc
{
  "schemaVersion": 2,
  "list":     { "title": "Project Tracker", "internalName": "ProjectTracker", "description": "" },
  "branding": { "title": "", "subtitle": "", "icon": "🦏", "accent": "#0f6cbd", "itemLabel": "project", "itemLabelPlural": "projects" },
  "fields":   [ { "internalName": "Title", "displayName": "Project Name", "type": "Text", "required": true } ],
  "form":     { "surface": "panel", "size": "medium", "sections": [ { "id": "main", "columns": 2, "fields": [ { "field": "Title", "width": "full" } ] } ] },
  "grid":     { "columns": [ { "field": "Title", "sortable": true, "filterable": true } ], "defaultSort": { "field": "ID", "direction": "desc" } }
}
```

### Field types

`Text`, `Note`, `Choice`, `MultiChoice`, `DateTime` (`dateOnly`), `Number` (`numberFormat`: decimal | integer | currency | percent, `decimals`, `min`, `max`), `Boolean`, `User` (`allowMultiple`), `URL`.

System columns (`ID`, `Created`, `Modified`, `Author`, `Editor`) can be placed on the form (read-only) or in the grid without being declared in `fields`.

### Form

| Key | Values |
| --- | --- |
| `surface` | `panel` (side panel) or `dialog` |
| `size` | `medium`, `large`, `extraLarge` |
| `addTitle` / `editTitle` | templates; `{item}`, `{Title}`, `{Id}` |
| `sections[]` | `id`, `title`, `description`, `columns` (1–3), `collapsible`, `collapsedByDefault`, `fields[]` |
| `sections[].fields[]` | `field`, `width` (`full`, `half`, `third`, `twoThirds`), `label`, `placeholder`, `helpText`, `readOnly` |

### Grid

| Key | Values |
| --- | --- |
| `columns[]` | `field`, `label`, `width`, `maxWidth`, `sortable`, `filterable`, `render`, `align`, `badgeTones` |
| `render` | `auto`, `text`, `badge`, `tags`, `persona`, `date`, `datetime`, `number`, `currency`, `percent`, `link`, `check` |
| `badgeTones` | `{ "Complete": "success", "On Hold": "neutral" }` – tones: neutral, info, success, warning, danger, purple, teal |
| `defaultSort` | `{ "field": "DueDate", "direction": "asc" }` |
| `pageSize` | rows per page (server-side paging) |
| `density` | `comfortable` or `compact` |
| `rowClick` | `edit`, `view`, `none` |
| toggles | `allowSearch`, `allowFilter`, `allowSort`, `allowExport`, `allowAdd`, `allowEdit`, `allowDelete`, `allowColumnResize`, `showCount` |

Filtering and sorting run server-side through OData; the search box filters the loaded page client-side.

## Project layout

```
src/
  shared/
    types/RhinoConfig.ts        the schema
    config/                     helpers, normaliser (v1 → v2), validator, sample data
    services/                   ListSetupService (provision / import), GridDataService (CRUD), ConfigStore (Site Assets), PeopleService
    components/                 DataGrid, FilterBar, ItemForm, FieldEditor, PeoplePicker, cell renderers, RhinoList
  webparts/
    rhinoGrid/                  "Rhino List" web part
    rhinoDesigner/              "Rhino List Designer" web part
samples/project-tracker.json
```

## Build and deploy (PowerShell)

Requires Node 22.

```powershell
npm ci
npx gulp build
npx gulp bundle --ship
npx gulp package-solution --ship
# Upload sharepoint/solution/rhino-forms.sppkg to the App Catalog, then add "Rhino Forms" to the site.
```

One-shot deploy to a tenant (App Catalog upload, tenant deploy, site install, sample config, Designer and List pages):

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
./deploy/Deploy-RhinoForms.ps1 -SiteUrl https://<tenant>.sharepoint.com
```

Local workbench:

```powershell
npx gulp serve
```

Bump `version` in `package.json` only; the build syncs `config/package-solution.json`.

## Notes

- Dates render as dd/mm/yyyy and currency as AUD.
- Date-only columns are saved at local midday so time zones never shift the calendar day.
- Deleting an item sends it to the site recycle bin.
- The Designer keeps an autosaved draft in the browser; **Save to site** publishes it.
