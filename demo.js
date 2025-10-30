import { CustomTable } from './lib/custom-table.js';

const container = document.getElementById('tableContainer');
const exportBtn = document.getElementById('exportBtn');
const exportToExcelBtn = document.getElementById('exportToExcelBtn');
const importInput = document.getElementById('importInput');
const trimRowBtn = document.getElementById('trimRowBtn');
const trimColBtn = document.getElementById('trimColBtn');

const table = new CustomTable(container, { rows: 2, cols: 2});

exportBtn.addEventListener('click', () => {
  // Export in spreadsheet-style JSON (commercial.json-like)
  const blob = new Blob([JSON.stringify(table.toSpreadsheetJSON(), null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'table-data-spreadsheet.json';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});
exportToExcelBtn.addEventListener('click', () => {
  table.exportToExcel('table-data.xlsx');
});

importInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const text = await file.text();
    const json = JSON.parse(text);
    table.fromJSON(json);
  } catch (err) {
    alert('Invalid JSON file.');
    console.error(err);
  } finally {
    importInput.value = '';
  }
});
function exportToExcel() {
  table.exportToExcel('table-data.xlsx');
}
trimRowBtn.addEventListener('click', () => {
  table.trimRow();
});
trimColBtn.addEventListener('click', () => {
  table.trimCol();
});


// Expose globally for quick console experiments (optional)
window.ctable = table;
