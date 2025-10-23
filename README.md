# Custom Table (mklctable)

Lightweight, dependency‑free table/spreadsheet component with column letters (A, B, C…), inline editing, per‑cell/row/column styling, merges, and export. Now distributed on npm and consumable from vanilla JS or Vue.

## Features
- Editable grid (contenteditable TDs)
- Column headers: A, B, C… (Excel‑style)
- Add/remove rows and columns
- Style toolbar: text align, bold/italic, text/background color, column width
- Per‑cell, per‑row, and per‑column inline styles with precedence: default → column → row → cell
- Data format: Spreadsheet‑style JSON only (commercial.json‑like); legacy internal model has been removed from the public API
- Import/Export: spreadsheet‑style JSON with consolidated keys
- Merged cells: import, render (rowSpan/colSpan), and export
- Selection restore on import (activeCell/selection)
- Applying a style to a whole row/column now updates each individual cell’s style too

## Project layout
- `index.html` – demo page
- `demo.js` – demo wiring (load/save)
- `lib/custom-table.js` – main ES module (CustomTable class)
- `lib/vue-custom-table.js` – optional Vue 3 wrapper (Composition API)
- `lib/custom-table.css` – component styles
- `dist/custom-table.css` – re‑export of CSS for npm consumers (`import 'mklctable/dist/custom-table.css'`)
- `styles.css` – demo styles
- `sample-data.json` / `commercial.json` – example data

## Run locally
Because this uses ES modules, serve with a simple web server.

### Simple server (PowerShell)
```powershell
# Serve current directory at http://localhost:8000
# Requires Python installed and on PATH
python -m http.server 8000
```
Open http://localhost:8000 and navigate to `index.html`.

## Installation

Using npm:

```bash
npm i mklctable
```

CDN (for quick demos):

```html
<link rel="stylesheet" href="https://unpkg.com/mklctable/dist/custom-table.css" />
<script type="module">
  import { CustomTable } from 'https://unpkg.com/mklctable/lib/custom-table.js';
  // use as shown below
</script>
```

## Usage (vanilla JS)

Bundler (Vite/Webpack/etc.):

```ts
import 'mklctable/dist/custom-table.css';
import { CustomTable } from 'mklctable/lib/custom-table.js';

const el = document.getElementById('table');
const table = new CustomTable(el, { rows: 4, cols: 4 });

// Styling APIs
table.setCellStyle(0, 0, { textAlign: 'center', background: '#fff3cd' });
table.setRowStyle(1, { fontWeight: 'bold' });
table.setColumnStyle(2, { width: '180px', color: '#0d6efd' });

// Spreadsheet JSON I/O
const sheetJson = table.getModel();         // Spreadsheet JSON
table.setModel(sheetJson);                  // Spreadsheet JSON
```

Include Material Icons (optional, for toolbar icons):

```html
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet" />
```

### Public API
- `new CustomTable(container, { rows?: number, cols?: number })`
- `addRow()` / `removeRow(index)`
- `addColumn()` / `removeColumn(index)`
- `toJSON()` – returns Spreadsheet JSON (alias of `toSpreadsheetJSON()`)
- `toSpreadsheetJSON()` – export Spreadsheet JSON (commercial.json‑like)
- `exportToExcel(filename?: string)` – downloads .xlsx (requires SheetJS); falls back to CSV
- `exportToCSV(filename?: string)` – downloads CSV (values only)
- `fromJSON(obj)` – Spreadsheet JSON only
- `getModel()` / `setModel(sheetJson)` – Spreadsheet JSON only
- `setCellStyle(r, c, style)` / `getCellStyle(r, c)`
- `getEffectiveCellStyle(r, c)` – computed cascade: default → column → row → cell
- `setRowStyle(r, style)` / `getRowStyle(r)`
- `setColumnStyle(c, style)` / `getColumnStyle(c)`
- `destroy()`

## Data format (Spreadsheet JSON only)
This component now reads and writes a commercial.json‑like spreadsheet shape. The legacy internal model has been removed from the public API. Use `getModel()` / `setModel()` or `toJSON()` / `fromJSON()` with the format below.

## Spreadsheet JSON import/export
`fromJSON` accepts a spreadsheet‑style shape, and `toSpreadsheetJSON()` exports the same shape with consolidated style keys:

```jsonc
{
  "activeSheet": "Sheet1",
  "sheets": [
    {
      "name": "Sheet1",
      "columns": [ {"width": 160}, {"width": 120} ],
      "rows": [
        { "index": 0, "height": 24, "cells": [
          { "index": 0, "value": "Name", "style": { "bold": true, "textAlign": "center" } },
          { "index": 1, "value": "Age" }
        ]},
        { "index": 1, "cells": [
          { "index": 0, "value": "Alice", "style": { "color": "#333" } },
          { "index": 1, "value": 30 }
        ]}
      ],
      "defaultCellStyle": { "fontFamily": "Segoe UI", "fontSize": 13 },
      "mergedCells": [ "A1:B1" ],
      "activeCell": "A1",
      "selection": "A1:B1"
    }
  ],
  "rowHeight": 22,
  "columnWidth": 100
}
```

What’s imported/exported
- Values: from `value` | `text` | `displayText` | `v`
- Grid size: inferred from `rows[].index`, `cells[].index`, `columns.length`, and A1 refs (`activeCell`, `selection`, `mergedCells`)
- Column widths: `columns[*].width` (px); fallback to top‑level `columnWidth`
- Row heights: `rows[*].height` (px); fallback to top‑level `rowHeight`
- Default cell style: `defaultCellStyle` cascades to all cells (import) and is emitted from current defaults (export)
- Per‑cell styles: inline top‑level keys or nested under `cell.style`/`cell.s` are supported on import; export writes under `cell.style`
- Selection: activeCell/selection is imported (restores selection) and exported
- Merged cells: imported and rendered (rowSpan/colSpan) and exported

### Export to Excel (.xlsx)
This project can export directly to .xlsx when the SheetJS library is present; otherwise it will fall back to CSV.

Add SheetJS via CDN in your `index.html` before your script that calls `exportToExcel`:

```html
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
```

Usage:

```js
import { CustomTable } from './lib/custom-table.js';
const table = new CustomTable(document.getElementById('table'));
// ... populate or import data ...
table.exportToExcel('table.xlsx'); // uses merges, column widths, row heights
```

Style key mapping (examples)
- Horizontal align: `hAlign|textAlign` → `textAlign` (export uses `textAlign`)
- Vertical align: `vAlign|verticalAlign` → `verticalAlign` (`middle` CSS maps to `center` in export)
- Wrapping: `wrap|wordWrap|wrapText` → `whiteSpace: 'normal'|'nowrap'`
- Font: `fontFamily`, `fontSize` (number → `px`)
- Colors: `color|fontColor|foreColor` → `color`; `background|bgColor|backColor|backgroundColor|fillColor` → `background` (export uses `background`)
- Emphasis: `bold|fontWeight`, `italic|fontStyle`, `underline/strike` → `textDecoration`

Notes
- Default style is applied to each cell first (handles non‑inheritable CSS like `verticalAlign`)
- Then column style, then row style, then cell style (overrides)
- Applying a style to a selected row/column also writes per‑cell overrides so all cells match
- Import accepts inline style keys at the cell object level or nested in `style`

## Toolbar and selection
- Target label shows the selection (Cell A1, Column B, Row R3)
- Buttons: align left/center/right; bold/italic
- Color inputs: text color, background color
- Column‑only input: width (e.g., `120px`)
- Apply styles to the current selection; Clear to remove styles

## Development notes
- No build tooling required; plain ES modules + CSS
- Keep changes focused in `lib/custom-table.js` and `lib/custom-table.css`
- When adding new spreadsheet style keys, extend the mapper in `#fromSpreadsheetJSON`

## Roadmap / known limitations
- UI to create/clear merged regions interactively
- Additional style coverage: borders, number formats
- No persistence baked‑in; use `toJSON()`/`fromJSON()` with your storage

## Vue 3 usage (optional)

Import the wrapper and use it inside a Vue component. The wrapper manages the lifecycle of the underlying CustomTable instance and proxies Spreadsheet JSON.

```vue
<template>
  <div class="sheet">
    <CustomTableVue ref="table" :model="model" @ready="onReady" />
  </div>
</template>

<script setup>
import 'mklctable/dist/custom-table.css';
import { createCustomTableComponent } from 'mklctable/lib/vue-custom-table.js';
import { getCurrentInstance } from 'vue';

const app = getCurrentInstance().appContext.app;
const CustomTableVue = createCustomTableComponent(app.config.globalProperties.__VUE__ || {
  // If using standard Vue 3, you can pass the imported Vue module directly instead of this shim.
});

const model = $ref(null); // Spreadsheet JSON
function onReady(inst) {
  // Access native methods: inst.addRow(), inst.exportToExcel(), etc.
}
</script>
```

Note: You can also import the wrapper directly in a setup file and `app.component('CustomTableVue', createCustomTableComponent(Vue))`.

## Breaking changes

As of v1.0.0:
- Public API now uses Spreadsheet JSON only.
- `toJSON()` returns Spreadsheet JSON; `fromJSON()` accepts Spreadsheet JSON only.
- `getModel()`/`setModel()` operate on Spreadsheet JSON, not the old internal model.
- CSS for npm consumers is available at `mklctable/dist/custom-table.css`.

## Changelog

v1.0.0
- Vue 3 wrapper (`lib/vue-custom-table.js`)
- Material Icons toolbar and icon‑only controls
- Excel export includes cell styles, merges, column widths, row heights
- Public API switched to Spreadsheet JSON only
- CSS distributed under `dist/` for easy import

## License
MIT
