# Customizable Table (Developer Guide)

Lightweight, dependency‑free, Vanilla JS table/spreadsheet component with column letters (A, B, C…), inline editing, per‑cell/row/column styling, and JSON import/export. The demo is a single static page—no build step required.

## Features
- Editable grid (contenteditable TDs)
- Column headers: A, B, C… (Excel‑style)
- Add/remove rows and columns
- Style toolbar: text align, bold/italic, text/background color, column width
- Per‑cell, per‑row, and per‑column inline styles with precedence: default → column → row → cell
- Import: internal JSON format and spreadsheet‑style JSON, including inline and nested cell styles
- Export: internal JSON format (round‑trip) and spreadsheet‑style JSON (commercial.json‑like) using consolidated keys
- Merged cells: import, render (rowSpan/colSpan), and export
- Selection restore on import (activeCell/selection)
- Applying a style to a whole row/column now updates each individual cell’s style too

## Project layout
- `index.html` – demo page
- `demo.js` – demo wiring (load/save)
- `lib/custom-table.js` – main ES module (CustomTable class)
- `lib/custom-table.css` – component styles
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

## Using the component

Include CSS, create a container, and instantiate:

```html
<link rel="stylesheet" href="./lib/custom-table.css" />
<div id="table"></div>
<script type="module">
  import { CustomTable } from './lib/custom-table.js';
  const table = new CustomTable(document.getElementById('table'), { rows: 4, cols: 4 });

  // Rows/Columns
  table.addRow();
  table.addColumn();
  table.removeRow(0);
  table.removeColumn(0);

  // Data I/O (internal model)
  const model = table.toJSON();
  table.fromJSON(model);

  // Styling APIs
  table.setCellStyle(0, 0, { textAlign: 'center', background: '#fff3cd' });
  table.setRowStyle(1, { fontWeight: 'bold' });
  table.setColumnStyle(2, { width: '180px', color: '#0d6efd' });

  // Access model directly (normalized)
  console.log(table.getModel());
  // table.setModel({ rows: 2, cols: 2, data: [["A","B"],["C","D"]] });
  // table.destroy();
  window.ctable = table; // for debugging in console
<\/script>
```

### Public API
- `new CustomTable(container, { rows?: number, cols?: number })`
- `addRow()` / `removeRow(index)`
- `addColumn()` / `removeColumn(index)`
- `toJSON()` – returns the internal model (see below)
- `toSpreadsheetJSON()` – exports a commercial.json‑like spreadsheet JSON
- `exportToExcel(filename?: string)` – downloads an .xlsx (requires SheetJS). Falls back to CSV if XLSX is not available.
- `exportToCSV(filename?: string)` – downloads a CSV (values only)
- `fromJSON(obj)` – accepts internal model or spreadsheet JSON
- `getModel()` / `setModel(model)` – normalized internal model
- `setCellStyle(r, c, style)` / `getCellStyle(r, c)`
- `getEffectiveCellStyle(r, c)` – computed cascade: default → column → row → cell
- `setRowStyle(r, style)` / `getRowStyle(r)`
- `setColumnStyle(c, style)` / `getColumnStyle(c)`
- `destroy()`

## Internal JSON model
Used by `toJSON()`/`fromJSON()` and safe to persist.

```json
{
  "rows": 3,
  "cols": 3,
  "data": [["Name","Age","City"],["Alice","30","Seattle"],["Bob","26","Austin"]],
  "columnStyles": [ {"width":"160px"}, null, null ],
  "rowStyles": [ null, {"fontWeight":"bold"}, null ],
  "cellStyles": {
    "C1R1": { "textAlign": "center", "background": "#fff3cd" }
  }
}
```
Notes
- `cellStyles` keys are `C{col+1}R{row+1}` (0‑based indices in code; 1‑based in keys)
- Values are inline CSS style objects (camelCase keys preferred; arbitrary CSS property names also supported via `style.setProperty`)

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

## License
MIT
