import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';

const dir = mkdtempSync(join(tmpdir(), 'm365-users-xlsx-'));
let counter = 0;

/**
 * Write a workbook to a temp file and return its path.
 * @param {Record<string, Array<Array<unknown>>>} sheets - sheet name → rows (first row = headers)
 */
export async function writeWorkbook(sheets) {
  const workbook = new ExcelJS.Workbook();
  for (const [name, rows] of Object.entries(sheets)) {
    const sheet = workbook.addWorksheet(name);
    for (const row of rows) sheet.addRow(row);
  }
  const path = join(dir, `book-${counter++}.xlsx`);
  await workbook.xlsx.writeFile(path);
  return path;
}

/** dd/mm/yyyy for today shifted by `days`, as the employee file writes dates. */
export function dmy(days = 0) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getFullYear()}`;
}
