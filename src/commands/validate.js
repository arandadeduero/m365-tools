/**
 * validate command
 *
 * Reads the Excel file found in the real-excel/ folder and runs all validation
 * rules defined in src/validators/excel-rules.js, printing a colour-coded report.
 *
 * Usage:
 *   m365-users validate
 *
 * Exit codes:
 *   0  — no issues found
 *   1  — one or more validation issues found (or file/read error)
 */

import { readdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import chalk from 'chalk';
import { readExcel, findExcelFile } from '../utils/excel.js';
import { RULES } from '../validators/excel-rules.js';

// ── 1. Locate file ──────────────────────────────────────────────────────
async function getFilePath() {
  try {
    return await findExcelFile();
  } catch (err) {
    throw err;
  }
}

export async function validateCommand() {
  // ── 1. Locate file ──────────────────────────────────────────────────────
  let filePath;
  try {
    filePath = await getFilePath();
  } catch (err) {
    console.error(chalk.red(`\nError: ${err.message}`));
    process.exit(1);
  }

  console.log(chalk.gray(`\nReading: ${filePath}`));

  // ── 2. Parse ─────────────────────────────────────────────────────────────
  let headers, rows;
  try {
    ({ headers, rows } = await readExcel(filePath));
  } catch (err) {
    console.error(chalk.red(`\nFailed to read Excel file: ${err.message}`));
    process.exit(1);
  }

  console.log(chalk.gray(`Found ${rows.length} data row(s) with ${headers.length} column(s).\n`));

  if (rows.length === 0) {
    console.log(chalk.yellow('The file contains no data rows. Nothing to validate.'));
    process.exit(0);
  }

  // ── 3. Run rules ─────────────────────────────────────────────────────────
  const allIssues = []; // { rule, issue }

  for (const rule of RULES) {
    const issues = rule.validate(rows);
    for (const issue of issues) {
      allIssues.push({ rule, issue });
    }
  }

  // ── 4. Print report ───────────────────────────────────────────────────────
  if (allIssues.length === 0) {
    console.log(chalk.green('✔  No issues found — the file looks good!'));
    process.exit(0);
  }

  // Group issues by rule for a clean report
  const byRule = new Map();
  for (const { rule, issue } of allIssues) {
    if (!byRule.has(rule.id)) byRule.set(rule.id, { rule, issues: [] });
    byRule.get(rule.id).issues.push(issue);
  }

  for (const { rule, issues } of byRule.values()) {
    console.log(chalk.bold.yellow(`● Rule: ${rule.id}`));
    console.log(chalk.gray(`  ${rule.description}`));
    console.log('');

    for (const issue of issues) {
      const valueDisplay = issue.value === '' ? chalk.italic('(empty)') : chalk.cyan(`"${issue.value}"`);
      const trabajador = rows[issue.row - 1]?.['Trabajador'] ?? '';
      const trabajadorDisplay = trabajador ? chalk.green(trabajador.padEnd(35)) : chalk.gray('(unknown)'.padEnd(35));
      console.log(
        `  ${chalk.red('✖')} Row ${chalk.bold(String(issue.row).padStart(4))}  ` +
          `${trabajadorDisplay}  ` +
          `${chalk.magenta(issue.column.padEnd(20))}  ` +
          `${valueDisplay.padEnd(30)}  ` +
          chalk.white(issue.message),
      );
    }

    console.log('');
  }

  // Tally affected rows (a row can have multiple issues)
  const affectedRows = new Set(allIssues.map(({ issue }) => issue.row));
  console.log(
    chalk.red.bold(
      `✖  ${allIssues.length} issue(s) found across ${affectedRows.size} row(s).`,
    ),
  );

  process.exit(1);
}
