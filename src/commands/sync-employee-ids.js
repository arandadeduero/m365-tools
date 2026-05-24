/**
 * sync-employee-ids command
 *
 * Reads id_empleado and Tipo de empleado from the Excel file in real-csv/ and
 * writes them to the employeeId and employeeType fields of each matching M365 user.
 *
 * Matching is done by email (e_mail → userPrincipalName).
 * The Excel is the source of truth: values are always overwritten.
 *
 * Valid values for Tipo de empleado (employeeType): Funcionario, Laboral.
 * Rows with an invalid or empty Tipo de empleado are flagged as warnings but
 * employeeId is still synced for that row.
 *
 * Rows are skipped when:
 *   - e_mail is empty or not @arandadeduero.es  (can't match cloud user)
 *   - both id_empleado and Tipo de empleado are already in sync
 *   - no matching cloud user found
 *
 * Usage:
 *   m365-users sync-employee-ids
 *
 * Exit codes:
 *   0  — all patches applied (or nothing to do)
 *   1  — one or more failures, aborted, or validation warnings present
 */

import { join, resolve } from 'node:path';
import { readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import chalk from 'chalk';
import inquirer from 'inquirer';
import { readExcel } from '../utils/excel.js';
import { listAllUsers, updateUser } from '../graph.js';

const PACKAGE_ROOT = resolve(fileURLToPath(import.meta.url), '../../../');
const REAL_CSV_DIR = join(PACKAGE_ROOT, 'real-csv');
const REQUIRED_DOMAIN = '@arandadeduero.es';

/** Allowed values for Tipo de empleado → employeeType */
const VALID_EMPLOYEE_TYPES = ['Funcionario', 'Laboral'];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function findExcelFile() {
  let entries;
  try {
    entries = await readdir(REAL_CSV_DIR);
  } catch {
    throw new Error(
      `Cannot read directory: ${REAL_CSV_DIR}\n` +
        '  Make sure the real-csv/ folder exists and contains an .xlsx file.',
    );
  }
  const file = entries.find((n) => n.endsWith('.xlsx') && !n.startsWith('~$'));
  if (!file) throw new Error(`No .xlsx file found in: ${REAL_CSV_DIR}`);
  return join(REAL_CSV_DIR, file);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

export async function syncEmployeeIdsCommand() {
  // ── 1. Load Excel ──────────────────────────────────────────────────────────
  let filePath;
  try {
    filePath = await findExcelFile();
  } catch (err) {
    console.error(chalk.red(`\nError: ${err.message}`));
    process.exit(1);
  }

  console.log(chalk.gray(`\nReading: ${filePath}`));

  let rows;
  try {
    ({ rows } = await readExcel(filePath));
  } catch (err) {
    console.error(chalk.red(`\nFailed to read Excel file: ${err.message}`));
    process.exit(1);
  }

  console.log(chalk.gray(`Excel: ${rows.length} row(s) found.\n`));

  // ── 2. Fetch cloud users ───────────────────────────────────────────────────
  console.log(chalk.gray('Fetching users from Microsoft 365…'));

  let cloudUsers;
  try {
    cloudUsers = await listAllUsers({ onProgress: () => {} });
  } catch (err) {
    console.error(chalk.red(`\nFailed to fetch users: ${err.message}`));
    process.exit(1);
  }

  // Build lookup: upn (lowercase) → user object (includes employeeId, employeeType)
  const cloudMap = new Map(
    cloudUsers.map((u) => [u.userPrincipalName.toLowerCase(), u]),
  );

  console.log(chalk.gray(`Cloud: ${cloudMap.size} user(s) found.\n`));

  // ── 3. Build patch list ────────────────────────────────────────────────────
  // toPatch entries: { name, email, cloudId, patch, preview }
  //   patch   — the PATCH body to send (only fields that changed)
  //   preview — display info per field
  const toPatch        = [];
  const skipped        = []; // { name, email, reason }
  const typeWarnings   = []; // { name, email, value } — invalid Tipo de empleado

  for (const row of rows) {
    const name         = row['Trabajador'] ?? '';
    const email        = (row['e_mail'] ?? '').trim().toLowerCase();
    const employeeId   = (row['id_empleado'] ?? '').toString().trim();
    const rawType      = (row['Tipo de empleado'] ?? '').toString().trim();

    if (!email.endsWith(REQUIRED_DOMAIN)) {
      skipped.push({ name, email: email || '(vacío)', reason: 'email sin dominio @arandadeduero.es' });
      continue;
    }

    const cloudUser = cloudMap.get(email);
    if (!cloudUser) {
      skipped.push({ name, email, reason: 'no encontrado en la nube' });
      continue;
    }

    const patch = {};

    // employeeId — always sync if present and different
    const cloudEmployeeId = (cloudUser.employeeId ?? '').toString().trim();
    if (employeeId && cloudEmployeeId !== employeeId) {
      patch.employeeId = employeeId;
    }

    // employeeType — validate, then sync if different
    let resolvedType = null;
    if (rawType) {
      if (VALID_EMPLOYEE_TYPES.includes(rawType)) {
        resolvedType = rawType;
        const cloudEmployeeType = (cloudUser.employeeType ?? '').toString().trim();
        if (cloudEmployeeType !== resolvedType) {
          patch.employeeType = resolvedType;
        }
      } else {
        typeWarnings.push({ name, email, value: rawType });
      }
    }

    if (Object.keys(patch).length === 0) {
      // Already fully in sync — silent skip
      continue;
    }

    toPatch.push({
      name,
      email,
      cloudId: cloudUser.id,
      patch,
      // display values
      employeeId:        patch.employeeId      ?? null,
      cloudEmployeeId:   cloudEmployeeId       || null,
      employeeType:      patch.employeeType    ?? null,
      cloudEmployeeType: (cloudUser.employeeType ?? '').toString().trim() || null,
    });
  }

  // ── 4. Show type validation warnings ──────────────────────────────────────
  if (typeWarnings.length > 0) {
    console.log(chalk.red.bold(`⚠  ${typeWarnings.length} fila(s) con valor inválido en "Tipo de empleado":`));
    console.log(chalk.gray(`   Valores permitidos: ${VALID_EMPLOYEE_TYPES.join(', ')}\n`));
    for (const w of typeWarnings) {
      console.log(`   ${chalk.white(w.name.padEnd(40))}  ${chalk.gray(w.email.padEnd(46))}  ${chalk.red(`"${w.value}"`)}`);
    }
    console.log('');
  }

  // ── 5. Print preview ──────────────────────────────────────────────────────
  if (toPatch.length === 0) {
    if (typeWarnings.length > 0) {
      console.log(chalk.yellow('Sin cambios que aplicar (corrige los valores inválidos antes de continuar).'));
      process.exit(1);
    }
    console.log(chalk.green('✔  Todos los campos ya están sincronizados — nada que hacer.'));
    process.exit(0);
  }

  const NAME_W = 36;
  const EMAIL_W = 44;
  const VAL_W  = 14;

  console.log(chalk.bold.yellow(`● ${toPatch.length} usuario(s) con cambios pendientes:\n`));

  // Header
  console.log(
    chalk.bold(
      '  ' + 'Nombre'.padEnd(NAME_W) + '  ' +
      'UPN'.padEnd(EMAIL_W) + '  ' +
      'Campo'.padEnd(14) + '  ' +
      'Excel'.padEnd(VAL_W) + '  ' +
      'Cloud actual',
    ),
  );
  console.log(chalk.gray('  ' + '─'.repeat(NAME_W + EMAIL_W + VAL_W + 42)));

  for (const u of toPatch) {
    const nameStr  = u.name.slice(0, NAME_W - 1).padEnd(NAME_W);
    const emailStr = u.email.padEnd(EMAIL_W);

    if (u.employeeId !== null) {
      const cloudLabel = u.cloudEmployeeId ? chalk.yellow(u.cloudEmployeeId) : chalk.gray('(vacío)');
      console.log(
        `  ${chalk.white(nameStr)}  ${chalk.gray(emailStr)}  ` +
        `${'employeeId'.padEnd(14)}  ${chalk.cyan(u.employeeId.padEnd(VAL_W))}  ${cloudLabel}`,
      );
    }
    if (u.employeeType !== null) {
      const cloudTypeLabel = u.cloudEmployeeType ? chalk.yellow(u.cloudEmployeeType) : chalk.gray('(vacío)');
      console.log(
        `  ${chalk.white(nameStr)}  ${chalk.gray(emailStr)}  ` +
        `${'employeeType'.padEnd(14)}  ${chalk.cyan(u.employeeType.padEnd(VAL_W))}  ${cloudTypeLabel}`,
      );
    }
  }

  if (skipped.length > 0) {
    console.log(chalk.gray(`\n  (${skipped.length} fila(s) omitidas por: email inválido o usuario no encontrado)`));
  }

  // ── 6. Confirmation ────────────────────────────────────────────────────────
  console.log('');
  const { confirm } = await inquirer.prompt([{
    type: 'confirm',
    name: 'confirm',
    message: chalk.yellow(`¿Aplicar cambios para ${toPatch.length} usuario(s) en Microsoft 365?`),
    default: false,
  }]);

  if (!confirm) {
    console.log(chalk.gray('\nCancelado — sin cambios.\n'));
    process.exit(1);
  }

  // ── 7. Apply ───────────────────────────────────────────────────────────────
  console.log('');
  let ok = 0;
  let failed = 0;

  for (const u of toPatch) {
    const fields = Object.keys(u.patch).join(', ');
    process.stdout.write(
      `  ${chalk.cyan(u.email.padEnd(EMAIL_W))}  ${chalk.gray(`[${fields}]`).padEnd(30)}  `,
    );
    try {
      await updateUser(u.cloudId, u.patch);
      console.log(chalk.green('✔'));
      ok++;
    } catch (err) {
      console.log(chalk.red(`✖  ${err.message}`));
      failed++;
    }
  }

  // ── 8. Summary ─────────────────────────────────────────────────────────────
  console.log('');
  if (failed === 0 && typeWarnings.length === 0) {
    console.log(chalk.green.bold(`✔  ${ok} usuario(s) actualizados correctamente.`));
    process.exit(0);
  } else {
    if (ok > 0)           console.log(chalk.green(`  ${ok} correctos`));
    if (failed > 0)       console.log(chalk.red(`  ${failed} fallidos`));
    if (typeWarnings.length > 0)
      console.log(chalk.yellow(`  ${typeWarnings.length} fila(s) con tipo de empleado inválido (sin cambios aplicados para esas filas)`));
    process.exit(1);
  }
}
