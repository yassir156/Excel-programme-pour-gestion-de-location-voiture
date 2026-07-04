import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export function exportToExcel(filename, sheets) {
  const workbook = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const sheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, sheet, name.slice(0, 31));
  });
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  saveAs(new Blob([buffer], { type: 'application/octet-stream' }), `${filename}.xlsx`);
}

export function exportToCSV(filename, rows) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(sheet);
  saveAs(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), `${filename}.csv`);
}
