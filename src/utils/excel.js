/**
 * Excel reader utility.
 *
 * Reads the "ACTIVOS" worksheet of an .xlsx file and returns rows as an array of
 * plain objects, using the first non-empty row as column headers.
 *
 * Values are trimmed strings; empty/null cells become empty strings ''.
 */

import ExcelJS from 'exceljs';
import { readdir } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PACKAGE_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
export const REAL_EXCEL_DIR = join(PACKAGE_ROOT, 'real-excel');

/**
 * Find the .xlsx file in real-excel/ that is not an Office temp file
 * (temp files start with ~$).
 */
export async function findExcelFile() {
  let entries;
  try {
    entries = await readdir(REAL_EXCEL_DIR);
  } catch {
    throw new Error(
      `Cannot read directory: ${REAL_EXCEL_DIR}\n` +
        '  Make sure the real-excel/ folder exists and contains an .xlsx file.',
    );
  }

  const files = entries.filter(
    (name) => name.endsWith('.xlsx') && !name.startsWith('~$'),
  );

  if (files.length === 0) {
    throw new Error(
      `No .xlsx file found in: ${REAL_EXCEL_DIR}\n` +
        '  Add the employee Excel file to the real-excel/ folder.',
    );
  }

  if (files.length > 1) {
    throw new Error(
      `Too many .xlsx files found in: ${REAL_EXCEL_DIR}\n` +
        '  Only one file should exist in the real-excel/ folder.\n' +
        `  Found: ${files.join(', ')}`,
    );
  }

  return join(REAL_EXCEL_DIR, files[0]);
}

/**
 * Coerce an ExcelJS cell value to a plain trimmed string.
 * ExcelJS can return rich text objects, hyperlink objects, formula results, etc.
 */
export function cellToString(value) {
  if (value == null) return '';

  // Date instance — format as dd/mm/yyyy (local time)
  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, '0');
    const mm = String(value.getMonth() + 1).padStart(2, '0');
    const yyyy = value.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  // Hyperlink cell: { text: '...', hyperlink: '...' }
  // The text itself can also be a rich text object, so recurse.
  if (typeof value === 'object' && 'text' in value) {
    return cellToString(value.text);
  }

  // Rich text: { richText: [{ text: '...' }, ...] }
  if (typeof value === 'object' && 'richText' in value) {
    return value.richText.map((r) => r.text ?? '').join('').trim();
  }

  // Formula result: { formula: '...', result: ... }
  if (typeof value === 'object' && 'result' in value) {
    return cellToString(value.result);
  }

  return String(value).trim();
}

/**
 * @param {string} filePath  Absolute path to the .xlsx file.
 * @returns {Promise<{ headers: string[], rows: Record<string, string>[] }>}
 */
export async function readExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet = workbook.getWorksheet('ACTIVOS');
  if (!sheet) {
    throw new Error(`Worksheet "ACTIVOS" not found in: ${filePath}`);
  }

  const allRows = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    const values = row.values.slice(1); // row.values is 1-indexed; slice off the leading undefined
    allRows.push(values);
  });

  if (allRows.length === 0) {
    return { headers: [], rows: [] };
  }

  // First row is the header
  const headers = allRows[0].map((h) => cellToString(h));

  const rows = allRows.slice(1).map((rowValues) => {
    const obj = {};
    headers.forEach((header, i) => {
      const raw = rowValues[i];
      obj[header] = cellToString(raw);
    });
    return obj;
  });

  return { headers, rows };
}
