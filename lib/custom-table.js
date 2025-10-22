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

    fromJSON(obj) {
        this.model = this.#normalizeModel(obj);
        this.#render();
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
            cellStyles: {}
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
        return { rows: safeRows, cols: safeCols, data, columnStyles, rowStyles, cellStyles };
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
            label.textContent = `C${c + 1}`;
            // Select column when clicking label/head (but not buttons)
            headWrap.addEventListener('click', (e) => { if (!e.target.closest('button')) this.#setSelection({ type: 'col', c }); });
            const btnRemove = document.createElement('button');
            btnRemove.type = 'button';
            btnRemove.className = 'btn danger remove-col';
            btnRemove.title = 'Remove this column';
            btnRemove.textContent = '−';
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
            rowRemoveBtnLeft.className = 'btn danger';
            rowRemoveBtnLeft.textContent = '−';
            rowRemoveBtnLeft.title = 'Remove this row';
            rowRemoveBtnLeft.addEventListener('click', () => this.removeRow(r));
            // Select row when clicking the row header area (but not the remove button)
            rowHead.addEventListener('click', (e) => { if (!e.target.closest('button')) this.#setSelection({ type: 'row', r }); });
            rowHead.appendChild(rowLabel);
            rowHead.appendChild(rowRemoveBtnLeft);
            thRow.appendChild(rowHead);
            tr.appendChild(thRow);

            for (let c = 0; c < this.model.cols; c++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.contentEditable = 'true';
                td.dataset.r = String(r);
                td.dataset.c = String(c);
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
                td.textContent = '+ col';
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
        tdAddRow.textContent = '+ row';
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
        wrap.appendChild(table);
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
          <button type="button" class="btn" data-style-prop="textAlign" data-style-val="left" title="Align left">⟸</button>
          <button type="button" class="btn" data-style-prop="textAlign" data-style-val="center" title="Align center">⇔</button>
          <button type="button" class="btn" data-style-prop="textAlign" data-style-val="right" title="Align right">⟹</button>
        </div>
        <div class="ct-group">
          <button type="button" class="btn" data-style-prop="fontWeight" data-style-val="bold" title="Bold">B</button>
          <button type="button" class="btn" data-style-prop="fontStyle" data-style-val="italic" title="Italic"><em>I</em></button>
        </div>
        <div class="ct-group">
          <input class="ct-input ct-color" data-style-prop="color" type="color" title="Text color" />
          <input class="ct-input ct-color" data-style-prop="background" type="color" title="Background" />
        </div>
        <div class="ct-group" data-scope="col-only" title="Column width">
          <input class="ct-input ct-width" data-style-prop="width" type="text" placeholder="120px" />
        </div>
        <div class="ct-group">
          <button type="button" class="btn secondary" data-action="clear">Clear</button>
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
                target.textContent = `Cell C${sel.c + 1}R${sel.r + 1}`;
                if (colOnly) colOnly.classList.add('disabled');
                style = this.getCellStyle(sel.r, sel.c) || {};
            } else if (sel?.type === 'col') {
                target.textContent = `Column C${sel.c + 1}`;
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
            this.setColumnStyle(sel.c, next);
        } else if (sel.type === 'row') {
            // Ignore column-only props for rows
            if ('width' in patch) delete patch.width;
            const cur = this.getRowStyle(sel.r) || {};
            const next = { ...cur };
            for (const [k, v] of Object.entries(patch)) {
                if (v == null || v === '') delete next[k]; else next[k] = v;
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
        <button type="button" class="btn danger ct-style-close" aria-label="Close">×</button>
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

    #editColumnStyle(index) {
        const current = this.getColumnStyle(index) || {};
        const sampleHint = '\n// Example: { "textAlign": "center", "width": "120px", "background": "#08101e" }';
        let input = null;
        try { input = prompt(`Edit style for column C${index + 1} as JSON:` + sampleHint, JSON.stringify(current, null, 2)); }
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
