// ES module library for reusable customizable table
// Public API: new CustomTable(container, options?), methods: addRow, addColumn, removeRow, removeColumn,
// toJSON(), fromJSON(obj), getModel(), setModel(model), destroy()

export class CustomTable {
    constructor(container, options = {}) {
        if (!container) throw new Error('CustomTable: container is required');
        this.container = container;
        const rows = options.rows ?? 2;
        const cols = options.cols ?? 2;
        this.model = this.#createEmptyModel(rows, cols);
        this._styleEditIndex = null;
        this._wrapEl = null;
        this._panelEl = null;
        this._selection = null; // { type: 'cell'|'col', r?, c? }
        this._defaultCellStyle = null; // spreadsheet default cell style applied to wrapper (font, etc.)
        this.#render();
    }

    // Public methods
    addRow() {
        const newRow = Array.from({ length: this.model.cols }, () => '');
        this.model.data.push(newRow);
        this.model.rows += 1;
        this.#render();
    }

    addColumn() {
        for (let r = 0; r < this.model.rows; r++) this.model.data[r].push('');
        this.model.cols += 1;
        if (!Array.isArray(this.model.columnStyles)) this.model.columnStyles = [];
        this.model.columnStyles.push(null);
        this.#render();
    }

    removeRow(index) {
        if (this.model.rows <= 1) return;
        if (index < 0 || index >= this.model.rows) return;
        this.model.data.splice(index, 1);
        this.model.rows -= 1;
        if (Array.isArray(this.model.rowStyles)) {
            this.model.rowStyles.splice(index, 1);
        }
        // Reindex or remove affected cell styles
        if (this.model.cellStyles) {
            this.model.cellStyles = this.#reindexCellStylesAfterRemoveRow(this.model.cellStyles, index);
        }
        // Adjust merged cells after row removal
        if (Array.isArray(this.model.mergedCells)) {
            this.model.mergedCells = this.#reindexMergesAfterRemoveRow(this.model.mergedCells, index);
        }
        this.#render();
    }

    removeColumn(index) {
        if (this.model.cols <= 1) return;
        if (index < 0 || index >= this.model.cols) return;
        for (let r = 0; r < this.model.rows; r++) this.model.data[r].splice(index, 1);
        this.model.cols -= 1;
        if (Array.isArray(this.model.columnStyles)) {
            this.model.columnStyles.splice(index, 1);
        }
        if (this.model.cellStyles) {
            this.model.cellStyles = this.#reindexCellStylesAfterRemoveCol(this.model.cellStyles, index);
        }
        if (Array.isArray(this.model.mergedCells)) {
            this.model.mergedCells = this.#reindexMergesAfterRemoveCol(this.model.mergedCells, index);
        }
        this.#render();
    }

    toJSON() {
        const out = structuredClone(this.model);
        if (out && out.cellStyles && typeof out.cellStyles === 'object') {
            const wrapped = {};
            for (const [k, v] of Object.entries(out.cellStyles)) {
                if (v && typeof v === 'object') wrapped[k] = { style: v };
            }
            out.cellStyles = wrapped;
        }
        return out;
    }

    // Export current table to a spreadsheet-style JSON (commercial.json-like)
    toSpreadsheetJSON() {
        const DEFAULT_COL_WIDTH = 64;
        const DEFAULT_ROW_HEIGHT = 21;

        // Columns: map width from columnStyles
        const columns = Array.from({ length: this.model.cols }, (_, c) => {
            const cs = this.model.columnStyles?.[c] || null;
            const width = this.#pxToNumber(cs?.width);
            return width != null ? { width } : {};
        });

        // Rows: sparse by default; include when value or style override exists
        const rows = [];
        for (let r = 0; r < this.model.rows; r++) {
            const rowStyle = this.model.rowStyles?.[r] || null;
            const height = this.#pxToNumber(rowStyle?.height);
            const cells = [];
            for (let c = 0; c < this.model.cols; c++) {
                const v = this.model.data[r][c] ?? '';
                const ck = this.#cellKey(r, c);
                const cellStyle = this.model.cellStyles?.[ck] || null;
                // Compute style overrides vs default for export
                const defaultStyle = this._defaultCellStyle || {};
                const colStyle = this.model.columnStyles?.[c] || null;
                const rowSty = rowStyle || null;
                const overrides = {};
                const layers = [colStyle, rowSty, cellStyle];
                for (const layer of layers) {
                    if (!layer || typeof layer !== 'object') continue;
                    for (const [k, val] of Object.entries(layer)) {
                        if (val == null || val === '') continue;
                        const dv = defaultStyle && k in defaultStyle ? defaultStyle[k] : undefined;
                        if (dv !== val) overrides[k] = val;
                    }
                }
                const styled = this.#mapToSpreadsheetStyle(overrides);
                const hasContent = typeof v === 'string' ? v.length > 0 : true;
                if (hasContent || (styled && Object.keys(styled).length)) {
                    const cell = { index: c };
                    if (hasContent) cell.value = v;
                    if (styled && Object.keys(styled).length) cell.style = styled;
                    cells.push(cell);
                }
            }
            if (cells.length > 0 || height != null) {
                const rowEntry = { index: r };
                if (height != null) rowEntry.height = height;
                rowEntry.cells = cells;
                rows.push(rowEntry);
            }
        }

        // Default cell style (if any) mapped back to spreadsheet keys
        const defaultCellStyle = this.#mapToSpreadsheetStyle(this._defaultCellStyle) || undefined;

        // Selection/activeCell
        let activeCell = 'A1';
        if (this._selection?.type === 'cell') {
            activeCell = `${this.#colIndexToLabel(this._selection.c)}${this._selection.r + 1}`;
        } else if (this._selection?.type === 'col') {
            activeCell = `${this.#colIndexToLabel(this._selection.c)}1`;
        } else if (this._selection?.type === 'row') {
            activeCell = `A${this._selection.r + 1}`;
        }

        const sheetName = 'Sheet1';
        const sheet = {
            name: sheetName,
            columns,
            rows,
            selection: activeCell,
            activeCell,
            frozenRows: 0,
            frozenColumns: 0,
            mergedCells: Array.isArray(this.model.mergedCells)
                ? this.model.mergedCells.map(m => this.#rangeToA1(m.start, m.end))
                : [],
            hyperlinks: [],
            defaultCellStyle,
            drawings: []
        };

        const out = {
            activeSheet: sheetName,
            sheets: [sheet],
            names: [],
            columnWidth: DEFAULT_COL_WIDTH,
            rowHeight: DEFAULT_ROW_HEIGHT,
            images: {}
        };
        return out;
    }

    // Export to Excel (.xlsx). Requires SheetJS (XLSX) to be loaded on the page.
    // If XLSX is not available, falls back to CSV download.
    exportToExcel(filename = 'table.xlsx') {
        try {
            // Prefer XLSX if available
            const XLSX = (typeof window !== 'undefined' && window.XLSX) ? window.XLSX : null;
            if (!XLSX) {
                console.warn('XLSX library not found. Falling back to CSV export.');
                this.exportToCSV(filename.replace(/\.xlsx$/i, '.csv'));
                return;
            }

            // Build AoA (Array of Arrays) of cell objects with value and style.
            // Covered cells in merges are blanked out.
            const rows = this.model.rows;
            const cols = this.model.cols;
            const aoa = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ v: '' })));
            const { coveredSet } = this.#buildMergeMaps(this.model?.mergedCells || []);
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const k = this.#posKey(r, c);
                    if (coveredSet.has(k)) {
                        aoa[r][c] = { v: '' };
                        continue;
                    }
                    const v = this.model.data[r][c] ?? '';
                    const eff = this.getEffectiveCellStyle(r, c) || null;
                    const style = this.#mapToXlsxCellStyle(eff);
                    // If style mapping exists, attach as cell.s; otherwise keep value only
                    aoa[r][c] = style ? { v, s: style } : { v };
                }
            }

            const ws = XLSX.utils.aoa_to_sheet(aoa);

            // Apply merges
            if (Array.isArray(this.model.mergedCells) && this.model.mergedCells.length) {
                ws['!merges'] = this.model.mergedCells.map(m => ({ s: { r: Math.min(m.start.r, m.end.r), c: Math.min(m.start.c, m.end.c) }, e: { r: Math.max(m.start.r, m.end.r), c: Math.max(m.start.c, m.end.c) } }));
            }

            // Column widths (pixels)
            if (Array.isArray(this.model.columnStyles)) {
                ws['!cols'] = this.model.columnStyles.map(cs => {
                    const wpx = this.#pxToNumber(cs?.width);
                    return wpx != null ? { wpx } : undefined;
                });
            }

            // Row heights (pixels)
            if (Array.isArray(this.model.rowStyles)) {
                ws['!rows'] = this.model.rowStyles.map(rs => {
                    const hpx = this.#pxToNumber(rs?.height);
                    return hpx != null ? { hpx } : undefined;
                });
            }

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            XLSX.writeFile(wb, filename);
        } catch (err) {
            console.error('Failed to export to Excel. Falling back to CSV.', err);
            this.exportToCSV(filename.replace(/\.xlsx$/i, '.csv'));
        }
    }

    // Export as CSV (Excel-compatible). Preserves values only.
    exportToCSV(filename = 'table.csv') {
        const rows = this.model.rows;
        const cols = this.model.cols;
        // Use merge map to blank covered cells for a cleaner CSV
        const { coveredSet } = this.#buildMergeMaps(this.model?.mergedCells || []);
        const esc = (v) => {
            const s = (v == null) ? '' : String(v);
            const needs = /[",\n\r]/.test(s);
            const doubled = s.replace(/"/g, '""');
            return needs ? `"${doubled}"` : doubled;
        };
        const lines = [];
        for (let r = 0; r < rows; r++) {
            const fields = [];
            for (let c = 0; c < cols; c++) {
                const k = this.#posKey(r, c);
                const v = coveredSet.has(k) ? '' : (this.model.data[r][c] ?? '');
                fields.push(esc(v));
            }
            lines.push(fields.join(','));
        }
        const bom = '\ufeff'; // ensure Excel opens as UTF-8
        const blob = new Blob([bom + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
        this.#downloadBlob(filename, blob);
    }

    #downloadBlob(filename, blob) {
        try {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 0);
        } catch (e) {
            console.error('Download failed', e);
        }
    }

    fromJSON(obj) {
        if (obj && typeof obj === 'object' && Array.isArray(obj.sheets)) {
            this.#fromSpreadsheetJSON(obj);
        } else {
            // Reset any spreadsheet-level default style so it doesn't leak between loads
            this._defaultCellStyle = null;
            this.model = this.#normalizeModel(obj);
            this.#render();
        }
    }

    getModel() { return this.model; }
    setModel(m) { this.model = this.#normalizeModel(m); this.#render(); }

    // Per-cell style APIs (r,c are 0-based)
    setCellStyle(r, c, styleObj) {
        if (r < 0 || r >= this.model.rows) return;
        if (c < 0 || c >= this.model.cols) return;
        if (!this.model.cellStyles) this.model.cellStyles = {};
        const key = this.#cellKey(r, c);
        const sanitized = this.#sanitizeStyle(styleObj);
        if (sanitized) this.model.cellStyles[key] = sanitized; else delete this.model.cellStyles[key];
        this.#render();
    }

    getCellStyle(r, c) {
        if (r < 0 || r >= this.model.rows) return null;
        if (c < 0 || c >= this.model.cols) return null;
        const key = this.#cellKey(r, c);
        return (this.model.cellStyles && this.model.cellStyles[key]) || null;
    }

    // Effective cell style = spreadsheet default + column + row + cell (later overrides earlier)
    getEffectiveCellStyle(r, c) {
        if (r < 0 || r >= this.model.rows) return null;
        if (c < 0 || c >= this.model.cols) return null;
        const key = this.#cellKey(r, c);
        const out = {};
        const layers = [this._defaultCellStyle, this.model.columnStyles?.[c], this.model.rowStyles?.[r], this.model.cellStyles?.[key]];
        for (const layer of layers) {
            if (!layer || typeof layer !== 'object') continue;
            for (const [k, v] of Object.entries(layer)) {
                if (v == null || v === '') continue;
                out[k] = v;
            }
        }
        return this.#sanitizeStyle(out);
    }

    // Column style APIs
    setColumnStyle(index, styleObj) {
        if (index < 0 || index >= this.model.cols) return;
        const sanitized = this.#sanitizeStyle(styleObj);
        if (!Array.isArray(this.model.columnStyles)) this.model.columnStyles = Array.from({ length: this.model.cols }, () => null);
        this.model.columnStyles[index] = sanitized;
        this.#render();
    }

    getColumnStyle(index) {
        if (index < 0 || index >= this.model.cols) return null;
        return (this.model.columnStyles && this.model.columnStyles[index]) || null;
    }

    // Row style APIs
    setRowStyle(index, styleObj) {
        if (index < 0 || index >= this.model.rows) return;
        const sanitized = this.#sanitizeStyle(styleObj);
        if (!Array.isArray(this.model.rowStyles)) this.model.rowStyles = Array.from({ length: this.model.rows }, () => null);
        this.model.rowStyles[index] = sanitized;
        this.#render();
    }

    getRowStyle(index) {
        if (index < 0 || index >= this.model.rows) return null;
        return (this.model.rowStyles && this.model.rowStyles[index]) || null;
    }

    destroy() {
        this.container.innerHTML = '';
    }

    // Private helpers
    #createEmptyModel(rows, cols) {
        return {
            rows,
            cols,
            data: Array.from({ length: rows }, () => Array.from({ length: cols }, () => '')),
            columnStyles: Array.from({ length: cols }, () => null),
            rowStyles: Array.from({ length: rows }, () => null),
            cellStyles: {},
            mergedCells: []
        };
    }

    #normalizeModel(m) {
        const DEFAULT_ROWS = 2, DEFAULT_COLS = 2;
        const rows = Math.max(0, Number(m?.rows ?? (m?.data?.length ?? DEFAULT_ROWS)));
        const cols = Math.max(0, Number(m?.cols ?? (Array.isArray(m?.data?.[0]) ? m.data[0].length : DEFAULT_COLS)));
        const safeRows = Number.isFinite(rows) ? rows : DEFAULT_ROWS;
        const safeCols = Number.isFinite(cols) ? cols : DEFAULT_COLS;
        const data = Array.from({ length: safeRows }, (__, r) => {
            const srcRow = Array.isArray(m?.data?.[r]) ? m.data[r] : [];
            return Array.from({ length: safeCols }, (___, c) => {
                const v = srcRow[c];
                return typeof v === 'string' ? v : (v == null ? '' : String(v));
            });
        });
        const srcStyles = Array.isArray(m?.columnStyles) ? m.columnStyles : [];
        const columnStyles = Array.from({ length: safeCols }, (_, c) => this.#sanitizeStyle(srcStyles[c]));
        const srcRowStyles = Array.isArray(m?.rowStyles) ? m.rowStyles : [];
        const rowStyles = Array.from({ length: safeRows }, (_, r) => this.#sanitizeStyle(srcRowStyles[r]));
        const srcCell = (m && typeof m.cellStyles === 'object' && !Array.isArray(m.cellStyles)) ? m.cellStyles : {};
        const cellStyles = {};
        for (const [k, v] of Object.entries(srcCell)) {
            const parsed = this.#parseCellKey(k);
            if (!parsed) continue;
            const { r, c } = parsed;
            if (r < safeRows && c < safeCols) {
                const sk = this.#cellKey(r, c);
                const raw = (v && typeof v === 'object' && 'style' in v) ? v.style : v;
                const sv = this.#sanitizeStyle(raw);
                if (sv) cellStyles[sk] = sv;
            }
        }
        const mergedCells = this.#sanitizeMerges(m?.mergedCells, safeRows, safeCols);
        return { rows: safeRows, cols: safeCols, data, columnStyles, rowStyles, cellStyles, mergedCells };
    }

    #sanitizeStyle(style) {
        if (!style || typeof style !== 'object') return null;
        const out = {};
        let count = 0;
        for (const [k, v] of Object.entries(style)) {
            if (count > 200) break; // prevent extremely large objects
            if (v == null) continue;
            const key = String(k).trim();
            const val = String(v);
            if (!key) continue;
            out[key] = val;
            count++;
        }
        return Object.keys(out).length ? out : null;
    }

    #render() {
        const container = this.container;
        container.innerHTML = '';

        const wrap = document.createElement('div');
        wrap.className = 'ctable table-wrap';
        this._wrapEl = wrap;
        // Apply default cell style (e.g., font family/size) to wrapper so it cascades
        if (this._defaultCellStyle) {
            this.#applyStyleObject(wrap, this._defaultCellStyle);
        }

        // Toolbar (Excel-like)
        const toolbar = this.#buildToolbar();
        wrap.appendChild(toolbar);

        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        // Header row: corner cell, column headers, trailing add-col cell
        const trHead = document.createElement('tr');
        const thCorner = document.createElement('th');
        thCorner.className = 'corner-head';
        trHead.appendChild(thCorner);

        for (let c = 0; c < this.model.cols; c++) {
            const th = document.createElement('th');
            th.dataset.c = String(c);
            const headWrap = document.createElement('div');
            headWrap.className = 'col-head';
            const label = document.createElement('span');
            label.textContent = this.#colIndexToLabel(c);
            // Select column when clicking label/head (but not buttons)
            headWrap.addEventListener('click', (e) => { if (!e.target.closest('button')) this.#setSelection({ type: 'col', c }); });
            const btnRemove = document.createElement('button');
            btnRemove.type = 'button';
            btnRemove.className = 'icon-btn remove-col';
            btnRemove.title = 'Remove this column';
            btnRemove.innerHTML = '<span class="material-icons" aria-hidden="true">remove</span>';
            btnRemove.addEventListener('click', () => this.removeColumn(c));
            headWrap.appendChild(label);
            headWrap.appendChild(btnRemove);
            th.appendChild(headWrap);
            // Apply column style to header cell
            this.#applyStyleObject(th, this.model.columnStyles?.[c]);
            trHead.appendChild(th);
        }

        const thAdd = document.createElement('th');

        trHead.appendChild(thAdd);
        thead.appendChild(trHead);

        // Precompute merge maps for rendering
        const { topLeftMap: mergeTopLeftMap, coveredSet: mergeCoveredSet } = this.#buildMergeMaps(this.model?.mergedCells || []);

        // Body rows
        for (let r = 0; r < this.model.rows; r++) {
            const tr = document.createElement('tr');
            tr.dataset.r = String(r);

            const thRow = document.createElement('th');
            const rowHead = document.createElement('div');
            rowHead.className = 'row-head';
            const rowLabel = document.createElement('span');
            rowLabel.textContent = `R${r + 1}`;
            rowLabel.title = 'Select row';
            rowLabel.addEventListener('click', () => { this.#setSelection({ type: 'row', r }); });
            const rowRemoveBtnLeft = document.createElement('button');
            rowRemoveBtnLeft.type = 'button';
            rowRemoveBtnLeft.className = 'icon-btn';
            rowRemoveBtnLeft.innerHTML = '<span class="material-icons" aria-hidden="true">remove</span>';
            rowRemoveBtnLeft.title = 'Remove this row';
            rowRemoveBtnLeft.addEventListener('click', () => this.removeRow(r));
            // Select row when clicking the row header area (but not the remove button)
            rowHead.addEventListener('click', (e) => { if (!e.target.closest('button')) this.#setSelection({ type: 'row', r }); });
            rowHead.appendChild(rowLabel);
            rowHead.appendChild(rowRemoveBtnLeft);
            thRow.appendChild(rowHead);
            tr.appendChild(thRow);

            for (let c = 0; c < this.model.cols; c++) {
                const posKey = this.#posKey(r, c);
                // Skip covered cells that are inside a merged region (not top-left)
                if (mergeCoveredSet.has(posKey)) continue;
                const td = document.createElement('td');
                td.className = 'cell';
                td.contentEditable = 'true';
                td.dataset.r = String(r);
                td.dataset.c = String(c);
                // If this cell is the top-left of a merged region, set spans
                const span = mergeTopLeftMap.get(posKey);
                if (span) {
                    if (span.rowSpan > 1) td.rowSpan = span.rowSpan;
                    if (span.colSpan > 1) td.colSpan = span.colSpan;
                }
                const val = this.model.data[r][c];
                if (val) { td.textContent = val; } else { td.classList.add('placeholder'); td.textContent = '' }
                td.addEventListener('input', (e) => {
                    const rr = Number(e.currentTarget.dataset.r);
                    const cc = Number(e.currentTarget.dataset.c);
                    const value = e.currentTarget.textContent ?? '';
                    if (Number.isFinite(rr) && Number.isFinite(cc)) {
                        this.model.data[rr][cc] = value;
                    }
                });
                td.addEventListener('focus', (e) => { e.currentTarget.classList.remove('placeholder'); this.#setSelection({ type: 'cell', r, c }); });
                // Apply spreadsheet default cell style first so non-inheritable defaults (e.g., verticalAlign) take effect
                this.#applyStyleObject(td, this._defaultCellStyle);
                // Apply column then row style to each cell
                this.#applyStyleObject(td, this.model.columnStyles?.[c]);
                this.#applyStyleObject(td, this.model.rowStyles?.[r]);
                // Apply cell-specific style (overrides column/row)
                const ck = this.#cellKey(r, c);
                const cs = this.model.cellStyles?.[ck];
                this.#applyStyleObject(td, cs);
                tr.appendChild(td);
            }
            if (r == 0) {
                const td = document.createElement('td');
                td.className = 'add-col';
                td.innerHTML = '<span class="material-icons" aria-hidden="true">add</span>';
                td.title = 'Add column';
                td.tabIndex = 0;
                td.addEventListener('click', () => this.addColumn());
                td.rowSpan = this.model.rows;
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }

        const trAddRow = document.createElement('tr');
        const tdAddRow = document.createElement('td');
        tdAddRow.colSpan = this.model.cols + 1;
    tdAddRow.className = 'add-row';
    tdAddRow.innerHTML = '<span class="material-icons" aria-hidden="true">add</span>';
        tdAddRow.title = 'Add row';
        tdAddRow.tabIndex = 0;
        tdAddRow.addEventListener('click', () => this.addRow());
        trAddRow.appendChild(tdAddRow);
        tbody.appendChild(trAddRow);

        table.appendChild(thead);
        table.appendChild(tbody);
        const caption = document.createElement('caption');
        caption.textContent = `${this.model.rows} rows × ${this.model.cols} cols`;
        table.appendChild(caption);
        const ctTable = document.createElement('div');
        ctTable.className = 'ctTable';
        ctTable.style.maxWidth = '100%';
        ctTable.style.overflowX = 'auto';
        ctTable.appendChild(table);
        wrap.appendChild(ctTable);
        container.appendChild(wrap);

        // Ensure style editor panel exists (appended inside wrap so we can position absolutely)
        this.#ensureStylePanel();

        // Try to restore focus/selection after render
        this.#restoreSelectionFocus();
        // Re-apply selection highlight after render
        this.#applySelectionStyles();
    }

    #buildToolbar() {
        const tb = document.createElement('div');
        tb.className = 'ct-toolbar';
        tb.innerHTML = `
            <div class="ct-tools">
                <span class="ct-target" data-role="target">—</span>
                <div class="ct-group">
                    <button type="button" class="icon-btn" data-style-prop="textAlign" data-style-val="left" title="Align left"><span class="material-icons" aria-hidden="true">format_align_left</span></button>
                    <button type="button" class="icon-btn" data-style-prop="textAlign" data-style-val="center" title="Align center"><span class="material-icons" aria-hidden="true">format_align_center</span></button>
                    <button type="button" class="icon-btn" data-style-prop="textAlign" data-style-val="right" title="Align right"><span class="material-icons" aria-hidden="true">format_align_right</span></button>
                </div>
                <div class="ct-group">
                    <button type="button" class="icon-btn" data-style-prop="fontWeight" data-style-val="bold" title="Bold"><span class="material-icons" aria-hidden="true">format_bold</span></button>
                    <button type="button" class="icon-btn" data-style-prop="fontStyle" data-style-val="italic" title="Italic"><span class="material-icons" aria-hidden="true">format_italic</span></button>
                </div>
                <div class="ct-group">
                    <input class="ct-input ct-color" data-style-prop="color" type="color" title="Text color" />
                    <input class="ct-input ct-color" data-style-prop="background" type="color" title="Background" />
                </div>
                <div class="ct-group" data-scope="col-only" title="Column width">
                    <input class="ct-input ct-width" data-style-prop="width" type="text" placeholder="120px" />
                </div>
                <div class="ct-group">
                    <button type="button" class="icon-btn" data-action="clear" title="Clear formatting"><span class="material-icons" aria-hidden="true">backspace</span></button>
                </div>
            </div>
        `;

        // Event delegation for buttons
        tb.addEventListener('click', (e) => {
            const t = e.target.closest('button');
            if (!t) return;
            const action = t.getAttribute('data-action');
            if (action === 'clear') { this.#applyStyleToSelection('clear'); return; }
            const prop = t.getAttribute('data-style-prop');
            const val = t.getAttribute('data-style-val');
            if (prop) { this.#applyStyleToSelection({ [prop]: val }); }
        });

        // Inputs change
        tb.addEventListener('change', (e) => {
            const input = e.target.closest('.ct-input');
            if (!input) return;
            const prop = input.getAttribute('data-style-prop');
            if (!prop) return;
            let val = (input.value || '').trim();
            if (input.type === 'color') {
                // Always use #RRGGBB format
                if (/^#[0-9a-fA-F]{6}$/.test(val)) {
                    // valid
                } else {
                    val = '#000000';
                }
            }
            if (!val) { this.#applyStyleToSelection({ [prop]: null }); }
            else { this.#applyStyleToSelection({ [prop]: val }); }
        });

        return tb;
    }

    #setSelection(sel) {
        this._selection = sel;
        // Update toolbar target label and scope
        const tb = this._wrapEl?.querySelector('.ct-toolbar');
        if (tb) {
            const target = tb.querySelector('[data-role="target"]');
            const colOnly = tb.querySelector('[data-scope="col-only"]');
            let style = null;
            if (sel?.type === 'cell') {
                const colLabel = this.#colIndexToLabel(sel.c);
                target.textContent = `Cell ${colLabel}${sel.r + 1}`;
                if (colOnly) colOnly.classList.add('disabled');
                // Prefill toolbar with effective style (default + col + row + cell)
                style = this.getEffectiveCellStyle(sel.r, sel.c) || {};
            } else if (sel?.type === 'col') {
                const colLabel = this.#colIndexToLabel(sel.c);
                target.textContent = `Column ${colLabel}`;
                if (colOnly) colOnly.classList.remove('disabled');
                style = this.getColumnStyle(sel.c) || {};
            } else if (sel?.type === 'row') {
                target.textContent = `Row R${sel.r + 1}`;
                if (colOnly) colOnly.classList.add('disabled');
                style = this.getRowStyle(sel.r) || {};
            } else {
                target.textContent = '—';
                if (colOnly) colOnly.classList.add('disabled');
            }
            this.#syncToolbarFromStyle(style, sel?.type);
        }
        // Update selection highlight
        this.#applySelectionStyles();
    }

    #applySelectionStyles() {
        const wrap = this._wrapEl; if (!wrap) return;
        // Clear previous highlights
        wrap.querySelectorAll('tbody tr.is-row-selected').forEach(tr => tr.classList.remove('is-row-selected'));
        wrap.querySelectorAll('thead th.is-col-selected, tbody td.is-col-selected').forEach(el => el.classList.remove('is-col-selected'));
        const sel = this._selection;
        if (sel?.type === 'row') {
            const tr = wrap.querySelector(`tbody tr[data-r="${sel.r}"]`);
            if (tr) tr.classList.add('is-row-selected');
        } else if (sel?.type === 'col') {
            wrap.querySelectorAll(`tbody td[data-c="${sel.c}"]`).forEach(td => td.classList.add('is-col-selected'));
            const th = wrap.querySelector(`thead th[data-c="${sel.c}"]`);
            if (th) th.classList.add('is-col-selected');
        }
    }

    #restoreSelectionFocus() {
        const sel = this._selection; if (!sel) return;
        if (sel.type === 'cell') {
            const td = this._wrapEl?.querySelector(`td[data-r="${sel.r}"][data-c="${sel.c}"]`);
            if (td) td.focus();
        }
    }

    #applyStyleToSelection(patch) {
        const sel = this._selection; if (!sel) return;
        if (patch === 'clear') {
            if (sel.type === 'cell') this.setCellStyle(sel.r, sel.c, null);
            else if (sel.type === 'col') this.setColumnStyle(sel.c, null);
            else if (sel.type === 'row') this.setRowStyle(sel.r, null);
            return;
        }
        if (sel.type === 'cell') {
            // Ignore column-only props for cells
            if ('width' in patch) delete patch.width;
            const cur = this.getCellStyle(sel.r, sel.c) || {};
            const next = { ...cur };
            for (const [k, v] of Object.entries(patch)) {
                if (v == null || v === '') delete next[k]; else next[k] = v;
            }
            this.setCellStyle(sel.r, sel.c, next);
        } else if (sel.type === 'col') {
            const cur = this.getColumnStyle(sel.c) || {};
            const next = { ...cur };
            for (const [k, v] of Object.entries(patch)) {
                if (v == null || v === '') delete next[k]; else next[k] = v;
            }
            // Apply to all cells in the column as per-cell overrides (excluding column-only props like width)
            const perCellPatch = { ...patch };
            if ('width' in perCellPatch) delete perCellPatch.width;
            if (!this.model.cellStyles) this.model.cellStyles = {};
            for (let r = 0; r < this.model.rows; r++) {
                const key = this.#cellKey(r, sel.c);
                const curCell = this.model.cellStyles[key] || {};
                const nextCell = { ...curCell };
                for (const [k, v] of Object.entries(perCellPatch)) {
                    if (v == null || v === '') delete nextCell[k]; else nextCell[k] = v;
                }
                const sanitized = this.#sanitizeStyle(nextCell);
                if (sanitized) this.model.cellStyles[key] = sanitized; else delete this.model.cellStyles[key];
            }
            this.setColumnStyle(sel.c, next);
        } else if (sel.type === 'row') {
            // Ignore column-only props for rows
            const rowPatch = { ...patch };
            if ('width' in rowPatch) delete rowPatch.width;
            const cur = this.getRowStyle(sel.r) || {};
            const next = { ...cur };
            for (const [k, v] of Object.entries(rowPatch)) {
                if (v == null || v === '') delete next[k]; else next[k] = v;
            }
            // Apply the same style to all cells in the row as per-cell overrides
            if (!this.model.cellStyles) this.model.cellStyles = {};
            for (let c = 0; c < this.model.cols; c++) {
                const key = this.#cellKey(sel.r, c);
                const curCell = this.model.cellStyles[key] || {};
                const nextCell = { ...curCell };
                for (const [k, v] of Object.entries(rowPatch)) {
                    if (v == null || v === '') delete nextCell[k]; else nextCell[k] = v;
                }
                const sanitized = this.#sanitizeStyle(nextCell);
                if (sanitized) this.model.cellStyles[key] = sanitized; else delete this.model.cellStyles[key];
            }
            this.setRowStyle(sel.r, next);
        }
    }

    #syncToolbarFromStyle(styleObj, selType) {
        const tb = this._wrapEl?.querySelector('.ct-toolbar');
        if (!tb) return;
        // Set input values
        tb.querySelectorAll('.ct-input').forEach((el) => {
            const key = el.getAttribute('data-style-prop');
            if (!key) return;
            let v = styleObj && key in styleObj ? styleObj[key] : '';
            // If input is type color, ensure value is valid hex
            if (el.type === 'color') {
                if (/^#[0-9a-fA-F]{6}$/.test(v)) {
                    el.value = v;
                } else {
                    // Light theme defaults: text black, background white
                    el.value = key === 'background' ? '#ffffff' : '#000000';
                }
            } else {
                el.value = v || '';
            }
        });
        // Toggle active on buttons
        const btns = tb.querySelectorAll('button[data-style-prop][data-style-val]');
        btns.forEach((btn) => {
            const prop = btn.getAttribute('data-style-prop');
            const val = btn.getAttribute('data-style-val');
            const active = styleObj && styleObj[prop] === val;
            btn.classList.toggle('active', !!active);
        });
        // Disable/enable width control for cell selections
        const colOnly = tb.querySelector('[data-scope="col-only"]');
        if (colOnly) {
            if (selType === 'col') {
                colOnly.classList.remove('disabled');
            } else {
                colOnly.classList.add('disabled');
                const widthInput = colOnly.querySelector('.ct-width');
                if (widthInput) widthInput.value = styleObj?.width || '';
            }
        }
    }

    // ----- Spreadsheet JSON support and A1 helpers -----
    #fromSpreadsheetJSON(json) {
        const sheets = Array.isArray(json.sheets) ? json.sheets : [];
        const activeName = json.activeSheet || sheets[0]?.name;
        const sheet = sheets.find(s => s.name === activeName) || sheets[0];
        if (!sheet) return;

        // Compute dimensions from sparse row/cell indices and A1 references
        let maxRow = -1;
        let maxCol = -1;
        if (Array.isArray(sheet.rows)) {
            for (const r of sheet.rows) {
                const ri = Number(r?.index);
                if (Number.isFinite(ri)) maxRow = Math.max(maxRow, ri);
                if (Array.isArray(r?.cells)) {
                    for (const cell of r.cells) {
                        const ci = Number(cell?.index ?? cell?.col ?? cell?.c);
                        if (Number.isFinite(ci)) maxCol = Math.max(maxCol, ci);
                    }
                }
            }
        }
        if (Array.isArray(sheet.columns)) maxCol = Math.max(maxCol, sheet.columns.length - 1);

        const bumpByA1 = (a1) => {
            const p = this.#parseA1(a1);
            if (p) { maxRow = Math.max(maxRow, p.r); maxCol = Math.max(maxCol, p.c); }
        };
        const bumpByRange = (rng) => {
            const pr = this.#parseA1Range(rng);
            if (pr) { maxRow = Math.max(maxRow, pr.end.r); maxCol = Math.max(maxCol, pr.end.c); }
        };
        if (typeof sheet.activeCell === 'string' && sheet.activeCell) bumpByA1(sheet.activeCell);
        if (typeof sheet.selection === 'string' && sheet.selection) bumpByRange(sheet.selection);
        const parsedMerges = [];
        if (Array.isArray(sheet.mergedCells)) {
            sheet.mergedCells.forEach((rng) => {
                bumpByRange(rng);
                const pr = this.#parseA1Range(rng);
                if (pr) parsedMerges.push(pr);
            });
        }

        const rows = Math.max(2, maxRow + 1);
        const cols = Math.max(2, maxCol + 1);

        // Initialize data grid
        const data = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));

        // Fill values from rows[].cells
        if (Array.isArray(sheet.rows)) {
            for (const r of sheet.rows) {
                const ri = Number(r?.index);
                if (!Number.isFinite(ri) || !Array.isArray(r.cells)) continue;
                for (const cell of r.cells) {
                    const ci = Number(cell?.index ?? cell?.col ?? cell?.c);
                    if (!Number.isFinite(ci)) continue;
                    let v = cell?.value;
                    if (v == null) v = cell?.text;
                    if (v == null) v = cell?.displayText;
                    if (v == null) v = cell?.v;
                    if (v == null) v = '';
                    data[ri][ci] = typeof v === 'string' ? v : String(v);
                }
            }
        }

        // Column widths
        const columnStyles = Array.from({ length: cols }, (_, c) => {
            const col = Array.isArray(sheet.columns) ? sheet.columns[c] : undefined;
            const widthNum = (col && typeof col.width === 'number') ? col.width : (typeof json.columnWidth === 'number' ? json.columnWidth : undefined);
            return Number.isFinite(widthNum) ? { width: `${Math.round(widthNum)}px` } : null;
        });

        // Row heights: map by explicit row.index, fallback to global rowHeight
        const rowHeights = new Map();
        if (Array.isArray(sheet.rows)) {
            for (const r of sheet.rows) {
                const ri = Number(r?.index);
                const h = Number(r?.height);
                if (Number.isFinite(ri) && Number.isFinite(h)) rowHeights.set(ri, h);
            }
        }
        const defaultRowHeight = Number.isFinite(json.rowHeight) ? json.rowHeight : undefined;
        const rowStyles = Array.from({ length: rows }, (_, r) => {
            const h = rowHeights.has(r) ? rowHeights.get(r) : defaultRowHeight;
            return Number.isFinite(h) ? { height: `${Math.round(h)}px` } : null;
        });

        // Cell styles (optional common mappings)
        // Prepare default style mapping to cascade typical text styles
        const mapStyle = (style) => {
            if (!style || typeof style !== 'object') return null;
            const out = {};
            if (style.textAlign) out.textAlign = String(style.textAlign);
            if (style.hAlign) out.textAlign = String(style.hAlign);
            if (style.vAlign) {
                const v = String(style.vAlign).toLowerCase();
                out.verticalAlign = v === 'center' ? 'middle' : v; // CSS expects 'middle'
            }
            if (style.verticalAlign) {
                const v = String(style.verticalAlign).toLowerCase();
                out.verticalAlign = v === 'center' ? 'middle' : v; // handle 'verticalAlign' too
            }
            if (style.wrap === true || style.wordWrap === true || style.wrapText === true) out.whiteSpace = 'normal';
            if (style.wrap === false || style.wordWrap === false || style.wrapText === false) out.whiteSpace = 'nowrap';
            if (style.fontFamily) out.fontFamily = String(style.fontFamily);
            if (style.fontSize != null && style.fontSize !== '') {
                const n = Number(style.fontSize);
                out.fontSize = Number.isFinite(n) ? `${Math.round(n)}px` : String(style.fontSize);
            }
            if (style.color || style.fontColor || style.foreColor) out.color = String(style.color || style.fontColor || style.foreColor);
            if (style.background || style.bgColor || style.backColor || style.backgroundColor || style.fillColor) {
                out.background = String(style.background || style.bgColor || style.backColor || style.backgroundColor || style.fillColor);
            }
            if (style.fontWeight) out.fontWeight = String(style.fontWeight);
            if (style.bold === true) out.fontWeight = 'bold';
            if (style.fontStyle) out.fontStyle = String(style.fontStyle);
            if (style.italic === true) out.fontStyle = 'italic';
            if (style.underline === true && style.strike === true) out.textDecoration = 'underline line-through';
            else if (style.underline === true) out.textDecoration = 'underline';
            else if (style.strike === true) out.textDecoration = 'line-through';
            return Object.keys(out).length ? out : null;
        };
        const defaultMapped = mapStyle(sheet.defaultCellStyle || json.defaultCellStyle) || null;
        // Save default so it cascades via wrapper
        this._defaultCellStyle = defaultMapped;
        const cellStyles = {};
        if (Array.isArray(sheet.rows)) {
            for (const r of sheet.rows) {
                const ri = Number(r?.index);
                if (!Number.isFinite(ri) || !Array.isArray(r.cells)) continue;
                for (const cell of r.cells) {
                    const ci = Number(cell?.index ?? cell?.col ?? cell?.c);
                    if (!Number.isFinite(ci)) continue;
                    const nestedStyle = cell?.style || cell?.s || null;
                    const mappedNested = mapStyle(nestedStyle) || {};
                    const mappedInline = mapStyle(cell) || {}; // map top-level style keys if present
                    const combined = { ...mappedNested, ...mappedInline };
                    // Merge: default first then cell-specific overrides (nested+inline)
                    const merged = defaultMapped ? { ...defaultMapped, ...combined } : combined;
                    const sanitized = this.#sanitizeStyle(merged);
                    if (sanitized) {
                        const key = this.#cellKey(ri, ci);
                        cellStyles[key] = sanitized;
                    }
                }
            }
        }

        const mergedCells = this.#sanitizeMerges(parsedMerges, rows, cols);

        // Determine selection to restore: prefer sheet.selection (range), else activeCell
        let sel = null;
        const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
        if (typeof sheet.selection === 'string' && sheet.selection.trim()) {
            const rng = this.#parseA1Range(sheet.selection.trim());
            if (rng) {
                const sr = clamp(rng.start.r, 0, rows - 1);
                const er = clamp(rng.end.r, 0, rows - 1);
                const sc = clamp(rng.start.c, 0, cols - 1);
                const ec = clamp(rng.end.c, 0, cols - 1);
                // Full row selected: spans all columns on a single row
                if (sr === er && sc === 0 && ec === cols - 1) {
                    sel = { type: 'row', r: sr };
                }
                // Full column selected: spans all rows on a single column
                else if (sc === ec && sr === 0 && er === rows - 1) {
                    sel = { type: 'col', c: sc };
                }
                // Otherwise, treat as a cell at the range start
                else {
                    sel = { type: 'cell', r: sr, c: sc };
                }
            }
        }
        if (!sel && typeof sheet.activeCell === 'string' && sheet.activeCell.trim()) {
            const ac = this.#parseA1(sheet.activeCell.trim());
            if (ac) {
                sel = { type: 'cell', r: clamp(ac.r, 0, rows - 1), c: clamp(ac.c, 0, cols - 1) };
            }
        }

        this.model = { rows, cols, data, columnStyles, rowStyles, cellStyles, mergedCells };
        this._selection = sel;
        this.#render();
    }

    // Column index (0-based) -> column label (A, B, ..., Z, AA, AB, ...)
    #colIndexToLabel(index) {
        let n = index;
        let label = '';
        do {
            const rem = n % 26;
            label = String.fromCharCode(65 + rem) + label;
            n = Math.floor(n / 26) - 1;
        } while (n >= 0);
        return label;
    }

    // Column label (e.g., 'A', 'Z', 'AA') -> index (0-based)
    #colLabelToIndex(label) {
        let n = 0;
        const s = String(label || '').trim().toUpperCase();
        if (!s) return 0;
        for (let i = 0; i < s.length; i++) {
            n = n * 26 + (s.charCodeAt(i) - 64);
        }
        return n - 1;
    }

    // Parse A1 like 'B2' -> {r,c}
    #parseA1(a1) {
        if (typeof a1 !== 'string') return null;
        const m = /^([A-Za-z]+)(\d+)$/.exec(a1.trim());
        if (!m) return null;
        const c = this.#colLabelToIndex(m[1]);
        const r1 = Number(m[2]);
        if (!Number.isFinite(r1)) return null;
        return { r: r1 - 1, c };
    }

    // Parse A1 range 'B2:D5' -> {start:{r,c}, end:{r,c}}
    #parseA1Range(rng) {
        if (typeof rng !== 'string') return null;
        const parts = rng.split(':');
        if (parts.length === 1) {
            const p = this.#parseA1(parts[0]);
            return p ? { start: p, end: p } : null;
        }
        const p1 = this.#parseA1(parts[0]);
        const p2 = this.#parseA1(parts[1]);
        if (!p1 || !p2) return null;
        return { start: p1, end: p2 };
    }

    // Helpers for cell style keys: C{col+1}R{row+1}
    #cellKey(r, c) { return `C${c + 1}R${r + 1}`; }
    #parseCellKey(k) {
        if (typeof k !== 'string') return null;
        const s = k.trim();
        const m = /^c\s*(\d+)\s*r\s*(\d+)$/i.exec(s);
        if (!m) return null;
        const c1 = Number(m[1]);
        const r1 = Number(m[2]);
        if (!Number.isFinite(c1) || !Number.isFinite(r1)) return null;
        return { c: c1 - 1, r: r1 - 1 };
    }

    #reindexCellStylesAfterRemoveRow(map, removedRow) {
        const out = {};
        for (const [k, v] of Object.entries(map)) {
            const parsed = this.#parseCellKey(k);
            if (!parsed) continue;
            const { r, c } = parsed;
            if (r === removedRow) continue;
            const newR = r > removedRow ? r - 1 : r;
            const nk = this.#cellKey(newR, c);
            out[nk] = v;
        }
        return out;
    }

    #reindexCellStylesAfterRemoveCol(map, removedCol) {
        const out = {};
        for (const [k, v] of Object.entries(map)) {
            const parsed = this.#parseCellKey(k);
            if (!parsed) continue;
            const { r, c } = parsed;
            if (c === removedCol) continue;
            const newC = c > removedCol ? c - 1 : c;
            const nk = this.#cellKey(r, newC);
            out[nk] = v;
        }
        return out;
    }

    #ensureStylePanel() {
        const wrap = this._wrapEl;
        if (!wrap) return;
        // Remove if already exists (we recreate per render to avoid stale refs)
        const old = wrap.querySelector('.ct-style-panel');
        if (old) old.remove();
        const panel = document.createElement('div');
        panel.className = 'ct-style-panel';
        panel.innerHTML = `
            <div class="ct-style-header">
        <span class="ct-style-title">Style</span>
            <button type="button" class="icon-btn ct-style-close" aria-label="Close"><span class="material-icons" aria-hidden="true">close</span></button>
            </div>
      <div class="ct-style-body">
        <label>Text align
          <select class="ct-input" data-key="textAlign">
            <option value="">Default</option>
            <option value="left">Left</option>
            <option value="center">Center</option>
            <option value="right">Right</option>
          </select>
        </label>
        <label>Width
          <input class="ct-input" type="text" placeholder="e.g. 120px" data-key="width" />
        </label>
        <label>Background
          <input class="ct-input" type="color" data-key="background" />
        </label>
        <label>Text color
          <input class="ct-input" type="color" data-key="color" />
        </label>
        <label>Font weight
          <select class="ct-input" data-key="fontWeight">
            <option value="">Default</option>
            <option value="normal">Normal</option>
            <option value="bold">Bold</option>
            <option value="600">600</option>
            <option value="700">700</option>
          </select>
        </label>
        <label>Font style
          <select class="ct-input" data-key="fontStyle">
            <option value="">Default</option>
            <option value="normal">Normal</option>
            <option value="italic">Italic</option>
          </select>
        </label>
      </div>
      <div class="ct-style-actions">
        <button type="button" class="btn secondary ct-style-reset">Reset</button>
        <div class="spacer"></div>
        <button type="button" class="btn secondary ct-style-cancel">Cancel</button>
        <button type="button" class="btn ct-style-save">Save</button>
      </div>
    `;
        panel.style.display = 'none';
        wrap.appendChild(panel);
        this._panelEl = panel;

        // Wire buttons
        panel.querySelector('.ct-style-close')?.addEventListener('click', () => this.#closeStylePanel());
        panel.querySelector('.ct-style-cancel')?.addEventListener('click', () => this.#closeStylePanel());
        panel.querySelector('.ct-style-reset')?.addEventListener('click', () => {
            const target = this._styleEditTarget;
            if (target) {
                if (target.type === 'row') this.setRowStyle(target.index, null);
                else if (target.type === 'col') this.setColumnStyle(target.index, null);
            }
            this.#closeStylePanel();
        });
        panel.querySelector('.ct-style-save')?.addEventListener('click', () => {
            const inputs = panel.querySelectorAll('.ct-input');
            const style = {};
            inputs.forEach((el) => {
                const key = el.getAttribute('data-key');
                let val = (el.value || '').trim();
                if (el.type === 'color') {
                    if (/^#[0-9a-fA-F]{6}$/.test(val)) {
                        // valid
                    } else {
                        val = '#000000';
                    }
                }
                if (key && val) { style[key] = val; }
            });
            const target = this._styleEditTarget;
            if (target) {
                if (target.type === 'row') this.setRowStyle(target.index, style);
                else if (target.type === 'col') this.setColumnStyle(target.index, style);
            }
            this.#closeStylePanel();
        });

        // Outside click to close
        const onDocClick = (e) => {
            if (panel.style.display === 'none') return;
            if (e.target === panel || panel.contains(e.target)) return;
            this.#closeStylePanel();
        };
        // Recreate listener per render to ensure we can remove it later if needed
        document.addEventListener('mousedown', onDocClick, { once: true });
    }

    #openStylePanel(type, index, anchor) {
        const panel = this._panelEl; const wrap = this._wrapEl;
        if (!panel || !wrap) return;
        this._styleEditTarget = { type, index };
        // Prefill
        const current = type === 'row' ? (this.getRowStyle(index) || {}) : (this.getColumnStyle(index) || {});
        const titleEl = panel.querySelector('.ct-style-title');
        if (titleEl) titleEl.textContent = type === 'row' ? `Row R${index + 1} style` : `Column C${index + 1} style`;
        panel.querySelectorAll('.ct-input').forEach((el) => {
            const key = el.getAttribute('data-key');
            let v = current && key && current[key] ? current[key] : '';
            if (el.type === 'color') {
                if (/^#[0-9a-fA-F]{6}$/.test(v)) {
                    el.value = v;
                } else {
                    el.value = key === 'background' ? '#ffffff' : '#000000';
                }
            } else {
                el.value = v || '';
            }
        });

        // Position near anchor
        const wrapRect = wrap.getBoundingClientRect();
        const anchorRect = anchor.getBoundingClientRect();
        const left = Math.max(8, anchorRect.left - wrapRect.left + wrap.scrollLeft - 8);
        const top = anchorRect.bottom - wrapRect.top + wrap.scrollTop + 6;
        panel.style.left = `${left}px`;
        panel.style.top = `${top}px`;
        panel.style.display = 'block';
    }

    #closeStylePanel() {
        if (this._panelEl) { this._panelEl.style.display = 'none'; }
        this._styleEditTarget = null;
    }

    #applyStyleObject(el, styleObj) {
        if (!styleObj) return;
        for (const [key, value] of Object.entries(styleObj)) {
            try {
                // Try direct property (camelCase) first
                if (key in el.style) {
                    // @ts-ignore - dynamic style assignment
                    el.style[key] = value;
                } else {
                    el.style.setProperty(key, value);
                }
            } catch { }
        }
    }

    // ----- Merge helpers -----
    #posKey(r, c) { return `${r}:${c}`; }

    #buildMergeMaps(mergedCells) {
        const topLeftMap = new Map(); // key -> { rowSpan, colSpan }
        const coveredSet = new Set(); // keys of cells that are covered (not top-left)
        if (!Array.isArray(mergedCells)) return { topLeftMap, coveredSet };
        for (const rng of mergedCells) {
            if (!rng || !rng.start || !rng.end) continue;
            const sr = Math.min(rng.start.r, rng.end.r);
            const er = Math.max(rng.start.r, rng.end.r);
            const sc = Math.min(rng.start.c, rng.end.c);
            const ec = Math.max(rng.start.c, rng.end.c);
            const rowSpan = er - sr + 1;
            const colSpan = ec - sc + 1;
            if (rowSpan <= 1 && colSpan <= 1) continue; // ignore 1x1
            const topKey = this.#posKey(sr, sc);
            topLeftMap.set(topKey, { rowSpan, colSpan });
            for (let r = sr; r <= er; r++) {
                for (let c = sc; c <= ec; c++) {
                    const k = this.#posKey(r, c);
                    if (k === topKey) continue;
                    coveredSet.add(k);
                }
            }
        }
        return { topLeftMap, coveredSet };
    }

    #sanitizeMerges(merges, maxRows, maxCols) {
        if (!Array.isArray(merges)) return [];
        const out = [];
        for (const m of merges) {
            let rng = null;
            if (typeof m === 'string') {
                rng = this.#parseA1Range(m);
            } else if (m && typeof m === 'object' && m.start && m.end) {
                const sr = Number(m.start.r), sc = Number(m.start.c);
                const er = Number(m.end.r), ec = Number(m.end.c);
                if (Number.isFinite(sr) && Number.isFinite(sc) && Number.isFinite(er) && Number.isFinite(ec)) {
                    rng = { start: { r: sr, c: sc }, end: { r: er, c: ec } };
                }
            }
            if (!rng) continue;
            // normalize order
            let sr = Math.max(0, Math.min(rng.start.r, rng.end.r));
            let er = Math.max(0, Math.max(rng.start.r, rng.end.r));
            let sc = Math.max(0, Math.min(rng.start.c, rng.end.c));
            let ec = Math.max(0, Math.max(rng.start.c, rng.end.c));
            // clamp within grid
            sr = Math.min(sr, maxRows - 1);
            er = Math.min(er, maxRows - 1);
            sc = Math.min(sc, maxCols - 1);
            ec = Math.min(ec, maxCols - 1);
            if (er < sr || ec < sc) continue;
            if (sr === er && sc === ec) continue; // ignore 1x1
            out.push({ start: { r: sr, c: sc }, end: { r: er, c: ec } });
        }
        return out;
    }

    #reindexMergesAfterRemoveRow(merges, removedRow) {
        const out = [];
        for (const m of merges) {
            const sr = m.start.r, er = m.end.r;
            const sc = m.start.c, ec = m.end.c;
            if (removedRow >= sr && removedRow <= er) {
                // row intersects merge -> drop it
                continue;
            }
            const shift = removedRow < sr ? -1 : 0;
            const nm = {
                start: { r: sr + shift, c: sc },
                end: { r: er + shift, c: ec }
            };
            if (nm.start.r <= nm.end.r) out.push(nm);
        }
        return out;
    }

    #reindexMergesAfterRemoveCol(merges, removedCol) {
        const out = [];
        for (const m of merges) {
            const sr = m.start.r, er = m.end.r;
            const sc = m.start.c, ec = m.end.c;
            if (removedCol >= sc && removedCol <= ec) {
                // col intersects merge -> drop it
                continue;
            }
            const shift = removedCol < sc ? -1 : 0;
            const nm = {
                start: { r: sr, c: sc + shift },
                end: { r: er, c: ec + shift }
            };
            if (nm.start.c <= nm.end.c) out.push(nm);
        }
        return out;
    }

    #rangeToA1(start, end) {
        if (!start || !end) return 'A1:A1';
        const sr = Math.min(start.r, end.r);
        const er = Math.max(start.r, end.r);
        const sc = Math.min(start.c, end.c);
        const ec = Math.max(start.c, end.c);
        const a1 = `${this.#colIndexToLabel(sc)}${sr + 1}`;
        const b1 = `${this.#colIndexToLabel(ec)}${er + 1}`;
        return `${a1}:${b1}`;
    }

    // Map internal CSS-like style object to spreadsheet-style shape
    #mapToSpreadsheetStyle(style) {
        if (!style || typeof style !== 'object') return null;
        const out = {};
        // Horizontal alignment: use textAlign only
        if (style.textAlign) {
            const ta = String(style.textAlign);
            out.textAlign = ta;
        }
        if (style.verticalAlign) {
            const v = String(style.verticalAlign).toLowerCase();
            const vv = v === 'middle' ? 'center' : v;
            // Vertical alignment: use verticalAlign only
            out.verticalAlign = vv;
        }
        if (style.whiteSpace) {
            const ws = String(style.whiteSpace).toLowerCase();
            if (ws === 'normal') out.wrap = true;
            else if (ws === 'nowrap') out.wrap = false;
        }
        if (style.fontFamily) out.fontFamily = String(style.fontFamily);
        if (style.fontSize != null && style.fontSize !== '') {
            const n = this.#pxToNumber(style.fontSize);
            if (n != null) out.fontSize = n; else out.fontSize = String(style.fontSize);
        }
        if (style.color) out.color = String(style.color);
        // Background: use background only
        if (style.background || style.backgroundColor) {
            const bg = String(style.background || style.backgroundColor);
            out.background = bg;
        }
        if (style.fontWeight) {
            const fw = String(style.fontWeight).toLowerCase();
            if (fw === 'bold') out.bold = true;
            else if (!isNaN(Number(fw)) && Number(fw) >= 600) out.bold = true;
        }
        if (style.fontStyle) {
            const fs = String(style.fontStyle).toLowerCase();
            if (fs === 'italic') out.italic = true;
        }
        if (style.textDecoration) {
            const td = String(style.textDecoration).toLowerCase();
            if (td.includes('underline')) out.underline = true;
            if (td.includes('line-through')) out.strike = true;
        }
        return Object.keys(out).length ? out : null;
    }

    // Map internal CSS-like style to a SheetJS (XLSX) cell style object.
    // Note: Applying styles requires a SheetJS build that supports style writing. Community builds may ignore styles.
    #mapToXlsxCellStyle(style) {
        if (!style || typeof style !== 'object') return null;
        const out = {};

        // Alignment
        const alignment = {};
        if (style.textAlign) {
            const h = String(style.textAlign).toLowerCase();
            if (h === 'left' || h === 'center' || h === 'right' || h === 'justify') alignment.horizontal = h;
        }
        if (style.verticalAlign) {
            const v = String(style.verticalAlign).toLowerCase();
            alignment.vertical = (v === 'middle') ? 'center' : v;
        }
        if (style.whiteSpace) {
            const ws = String(style.whiteSpace).toLowerCase();
            if (ws === 'normal') alignment.wrapText = true;
            else if (ws === 'nowrap') alignment.wrapText = false;
        }
        if (Object.keys(alignment).length) out.alignment = alignment;

        // Font
        const font = {};
        if (style.fontFamily) font.name = String(style.fontFamily);
        if (style.fontSize != null && style.fontSize !== '') {
            const px = this.#pxToNumber(style.fontSize);
            // Excel expects points; approx 1px = 0.75pt
            if (px != null) font.sz = Math.round(px * 0.75 * 10) / 10;
        }
        if (style.fontWeight) {
            const fw = String(style.fontWeight).toLowerCase();
            if (fw === 'bold' || (!isNaN(Number(fw)) && Number(fw) >= 600)) font.bold = true;
        }
        if (style.fontStyle) {
            const fs = String(style.fontStyle).toLowerCase();
            if (fs === 'italic') font.italic = true;
        }
        if (style.textDecoration) {
            const td = String(style.textDecoration).toLowerCase();
            if (td.includes('underline')) font.underline = true;
            if (td.includes('line-through')) font.strike = true;
        }
        if (style.color) {
            const rgb = this.#cssColorToXlsxRGB(style.color);
            if (rgb) font.color = { rgb };
        }
        if (Object.keys(font).length) out.font = font;

        // Fill (background)
        if (style.background || style.backgroundColor) {
            const rgb = this.#cssColorToXlsxRGB(style.background || style.backgroundColor);
            if (rgb) out.fill = { patternType: 'solid', fgColor: { rgb } };
        }

        return Object.keys(out).length ? out : null;
    }

    // Convert common CSS color strings (#RRGGBB, rgb(), named) to XLSX RGB hex (RRGGBB)
    #cssColorToXlsxRGB(input) {
        if (!input) return null;
        const s = String(input).trim().toLowerCase();
        if (!s) return null;
        // hex #rgb or #rrggbb or #rrggbbaa
        const mhex = /^#([0-9a-f]{3,8})$/i.exec(s);
        if (mhex) {
            let h = mhex[1];
            if (h.length === 3) { // rgb -> rrggbb
                h = h.split('').map(ch => ch + ch).join('');
            } else if (h.length === 8) { // rrggbbaa -> rrggbb
                h = h.slice(0, 6);
            } else if (h.length !== 6) {
                // unsupported hex length
                return null;
            }
            return h.toUpperCase();
        }
        // rgb()/rgba()
        const mrgb = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([0-9.]+))?\s*\)$/i.exec(s);
        if (mrgb) {
            const r = Math.max(0, Math.min(255, Number(mrgb[1])));
            const g = Math.max(0, Math.min(255, Number(mrgb[2])));
            const b = Math.max(0, Math.min(255, Number(mrgb[3])));
            const toHex = (n) => n.toString(16).toUpperCase().padStart(2, '0');
            return `${toHex(r)}${toHex(g)}${toHex(b)}`;
        }
        // named colors (common subset)
        const NAMED = {
            black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', lime: '00FF00',
            blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', silver: 'C0C0C0',
            maroon: '800000', navy: '000080', teal: '008080', purple: '800080', fuchsia: 'FF00FF',
            aqua: '00FFFF', orange: 'FFA500'
        };
        if (s in NAMED) return NAMED[s];
        if (s === 'transparent') return null;
        return null;
    }

    // Convert CSS px value or number-like string to number; returns null if not parsable
    #pxToNumber(val) {
        if (val == null || val === '') return null;
        if (typeof val === 'number' && Number.isFinite(val)) return Math.round(val);
        const s = String(val).trim();
        const m = /^(-?\d+(?:\.\d+)?)\s*px?$/i.exec(s) || /^-?\d+(?:\.\d+)?$/i.exec(s);
        if (m) {
            const n = Number(m[1] ?? m[0]);
            return Number.isFinite(n) ? Math.round(n) : null;
        }
        return null;
    }

    #editColumnStyle(index) {
        const current = this.getColumnStyle(index) || {};
        const sampleHint = '\n// Example: { "textAlign": "center", "width": "120px", "background": "#08101e" }';
        let input = null;
        try { input = prompt(`Edit style for column ${this.#colIndexToLabel(index)} as JSON:` + sampleHint, JSON.stringify(current, null, 2)); }
        catch { /* ignore */ }
        if (input == null) return; // cancelled
        try {
            const obj = JSON.parse(input);
            this.setColumnStyle(index, obj);
        } catch (err) {
            alert('Invalid JSON for style.');
            // no re-render here; setColumnStyle renders on success
        }
    }
}
