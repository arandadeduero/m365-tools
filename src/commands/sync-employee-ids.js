/**
 * sync-employee-ids
 *
 * Reads id_empleado and Tipo de empleado from the Excel file and
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
 */

import chalk from 'chalk';
import { updateUser } from '../graph.js';
import { REQUIRED_DOMAIN } from '../constants.js';

/** Allowed values for Tipo de empleado → employeeType */
const VALID_EMPLOYEE_TYPES = ['Funcionario', 'Laboral'];

// ---------------------------------------------------------------------------
// Core logic — reusable, no prompts, no process.exit
// ---------------------------------------------------------------------------

/**
 * Sync employeeId and employeeType from Excel rows to M365.
 * Applies changes immediately without confirmation (designed to run as part of import).
 *
 * @param {Array} rows        - Raw Excel rows (with e_mail, id_empleado, Tipo de empleado, Trabajador)
 * @param {Map}   cloudMap    - Map<upn (lowercase), cloudUser object>
 * @returns {{ ok: number, failed: number, skipped: number, typeWarnings: number }}
 */
export async function syncEmployeeIds(rows, cloudMap) {
  const toPatch      = [];
  const skipped      = [];
  const typeWarnings = [];

  for (const row of rows) {
    const name       = row['Trabajador'] ?? '';
    const email      = (row['e_mail'] ?? '').trim().toLowerCase();
    const employeeId = (row['id_empleado'] ?? '').toString().trim();
    const rawType    = (row['Tipo de empleado'] ?? '').toString().trim();

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
    if (rawType) {
      if (VALID_EMPLOYEE_TYPES.includes(rawType)) {
        const cloudEmployeeType = (cloudUser.employeeType ?? '').toString().trim();
        if (cloudEmployeeType !== rawType) {
          patch.employeeType = rawType;
        }
      } else {
        typeWarnings.push({ name, email, value: rawType });
      }
    }

    if (Object.keys(patch).length === 0) continue;

    toPatch.push({ name, email, cloudId: cloudUser.id, patch });
  }

  // ── Type warnings ─────────────────────────────────────────────────────────
  if (typeWarnings.length > 0) {
    console.log(chalk.red.bold(`\n⚠  ${typeWarnings.length} fila(s) con valor inválido en "Tipo de empleado":`));
    console.log(chalk.gray(`   Valores permitidos: ${VALID_EMPLOYEE_TYPES.join(', ')}\n`));
    for (const w of typeWarnings) {
      console.log(`   ${chalk.white(w.name.padEnd(40))}  ${chalk.gray(w.email.padEnd(46))}  ${chalk.red(`"${w.value}"`)}`);
    }
  }

  // ── Nothing to do ─────────────────────────────────────────────────────────
  if (toPatch.length === 0) {
    console.log(chalk.green('  ✔  Employee IDs already in sync — nothing to do.'));
    return { ok: 0, failed: 0, skipped: skipped.length, typeWarnings: typeWarnings.length };
  }

  // ── Apply patches ─────────────────────────────────────────────────────────
  let ok = 0;
  let failed = 0;

  for (const u of toPatch) {
    const fields = Object.keys(u.patch).join(', ');
    process.stdout.write(
      `  ${chalk.cyan(u.email.padEnd(44))}  ${chalk.gray(`[${fields}]`).padEnd(30)}  `,
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

  // ── Summary line ──────────────────────────────────────────────────────────
  if (failed === 0) {
    console.log(chalk.green(`  ✔  ${ok} employee ID(s) synced.`));
  } else {
    console.log(chalk.yellow(`  ${ok} synced, `) + chalk.red(`${failed} failed.`));
  }

  return { ok, failed, skipped: skipped.length, typeWarnings: typeWarnings.length };
}
