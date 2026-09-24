/**
 * sync-check command
 *
 * Reads the employee Excel file from path and cross-references it with
 * Microsoft 365 to detect two categories of sync issues.
 *
 * Check 1 — Must exist in cloud:
 *   All employees with no (effective) Fecha de Baja should have an account in
 *   M365. Those with an @arandadeduero.es email are looked up by UPN; those
 *   without a domain email are flagged separately as a data gap.
 *
 * Check 2 — Must be disabled in cloud:
 *   Employees with a past/present Fecha de Baja and an @arandadeduero.es email
 *   must have accountEnabled = false in M365.
 *
 * Check 3 — Cloud-only accounts:
 *   Users that exist in M365 with an @arandadeduero.es email but do not appear
 *   in the Excel file at all (neither as active nor as left employees).
 *
 * --fix flag (Check 2 only):
 *   Prints the list of accounts to act on, asks for double confirmation, then
 *   for each account: disables it (accountEnabled = false) and removes all
 *   assigned licences.
 *
 * Usage:
 *   m365-users sync-check <file>
 *   m365-users sync-check <file> --disable-left-workers
 *
 * Exit codes:
 *   0  — everything is in sync (or all fixes applied successfully)
 *   1  — issues found / fix aborted / error
 */

import chalk from 'chalk';
import inquirer from 'inquirer';
import ora from 'ora';
import { readExcel } from '../utils/excel.js';
import { listAllUsers, updateUser, getUserLicenses, removeAllLicenses } from '../graph.js';
import { REQUIRED_DOMAIN } from '../constants.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function parseDMY(value) {
  if (!value) return null;
  const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;
  const [, dd, mm, yyyy] = match;
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function isEffectiveLeaving(fechaBaja) {
  const date = parseDMY(fechaBaja);
  if (!date) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return date <= today;
}

// ---------------------------------------------------------------------------
// Fix flow — disable accounts + remove licences
// ---------------------------------------------------------------------------

async function fixNotDisabled(notDisabled) {
  // ── Preview ───────────────────────────────────────────────────────────────
  console.log('');
  console.log(chalk.bold.red('  The following accounts will be DISABLED and UNLICENSED:\n'));

  const actionable = notDisabled; // all entries are now guaranteed to exist in cloud

  // Fetch licences for all actionable accounts upfront so we can show them in
  // the preview and build the summary before asking for confirmation.
  console.log(chalk.gray('  Fetching current licence assignments…\n'));
  const licenceMap = new Map(); // email -> [{ skuId, skuPartNumber }]
  for (const info of actionable) {
    try {
      const lics = await getUserLicenses(info.email);
      licenceMap.set(info.email, lics);
    } catch {
      licenceMap.set(info.email, []);
    }
  }

  // ── Per-account rows ──────────────────────────────────────────────────────
  for (const info of notDisabled) {
    const lics = licenceMap.get(info.email) ?? [];
    const licLabel = lics.length === 0
      ? chalk.gray('no licences')
      : lics.map((l) => chalk.magenta(l.skuPartNumber)).join(', ');
    console.log(
      `  ${chalk.red('►')} ${chalk.green(info.name.padEnd(40))}  ` +
      `${chalk.cyan(info.email.padEnd(45))}  ` +
      `Fecha de Baja: ${chalk.yellow(info.fechaBaja.padEnd(12))}  ` +
      `Licences: ${licLabel}`,
    );
  }

  if (actionable.length === 0) {
    console.log(chalk.yellow('\n  No actionable accounts.'));
    return;
  }

  // ── Licence summary ───────────────────────────────────────────────────────
  const licenceTotals = new Map(); // skuPartNumber -> count
  for (const lics of licenceMap.values()) {
    for (const l of lics) {
      licenceTotals.set(l.skuPartNumber, (licenceTotals.get(l.skuPartNumber) ?? 0) + 1);
    }
  }

  const totalLicenceAssignments = [...licenceTotals.values()].reduce((a, b) => a + b, 0);

  console.log('');
  console.log(chalk.bold('  Licence removal summary:'));
  if (licenceTotals.size === 0) {
    console.log(chalk.gray('    None of the affected accounts hold any licences.'));
  } else {
    for (const [sku, count] of [...licenceTotals.entries()].sort()) {
      console.log(`    ${chalk.magenta(sku.padEnd(40))}  ${chalk.white(`${count} assignment(s) will be removed`)}`);
    }
    console.log(chalk.gray(`\n    Total: ${totalLicenceAssignments} licence assignment(s) across ${actionable.length} account(s).`));
  }

  console.log('');
  console.log(chalk.yellow(`  ${actionable.length} account(s) will be disabled and unlicensed.`));

  // ── First confirmation ────────────────────────────────────────────────────
  const { confirm1 } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm1',
      message: chalk.yellow('Are you sure you want to disable and unlicense these accounts?'),
      default: false,
    },
  ]);

  if (!confirm1) {
    console.log(chalk.gray('\nAborted — no changes made.'));
    process.exit(1);
  }

  // ── Second confirmation ───────────────────────────────────────────────────
  const { confirm2 } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm2',
      message: chalk.red(`FINAL confirmation: disable and remove licences for ${actionable.length} account(s)?`),
      default: false,
    },
  ]);

  if (!confirm2) {
    console.log(chalk.gray('\nAborted — no changes made.'));
    process.exit(1);
  }

  // ── Apply ─────────────────────────────────────────────────────────────────
  console.log('');
  let successCount = 0;
  let errorCount = 0;

  for (const info of actionable) {
    process.stdout.write(`  ${chalk.cyan(info.email.padEnd(45))}  `);
    try {
      // Use the already-fetched licence list (licenceMap built during preview).
      // Licences are removed before disabling — Graph won't allow licence removal
      // on a disabled account in some tenants.
      const licenses = licenceMap.get(info.email) ?? [];

      // 1. Remove all licences
      if (licenses.length > 0) {
        await removeAllLicenses(info.email, licenses);
        process.stdout.write(chalk.gray(`removed ${licenses.length} licence(s)  `));
      } else {
        process.stdout.write(chalk.gray('no licences         '));
      }

      // 2. Disable account
      await updateUser(info.email, { accountEnabled: false });
      process.stdout.write(chalk.green('disabled  '));

      console.log(chalk.green('✔'));
      successCount++;
    } catch (err) {
      console.log(chalk.red(`✖  ${err.message}`));
      errorCount++;
    }
  }

  console.log('');
  if (errorCount === 0) {
    console.log(chalk.green.bold(`✔  ${successCount} account(s) disabled and unlicensed successfully.`));
    process.exit(0);
  } else {
    console.log(
      chalk.yellow(`  ${successCount} succeeded, `) + chalk.red(`${errorCount} failed.`),
    );
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

export async function syncCheckCommand(filePath, { fix = false } = {}) {
  console.log(chalk.gray(`\nReading: ${filePath}`));

  let rows, concejalesRows;
  try {
    ({ rows, concejalesRows } = await readExcel(filePath));
  } catch (err) {
    console.error(chalk.red(`\nFailed to read Excel file: ${err.message}`));
    process.exit(1);
  }

  // ── 2. Build lookup sets from Excel ───────────────────────────────────────
  const mustExist = new Map();      // upn -> { name, email }
  const mustBeDisabled = new Map(); // upn -> { name, email, fechaBaja }
  const noEmail = [];               // { name, email } — active, no lookable UPN
  const allExcelEmails = new Set(); // All emails from both ACTIVOS and CONCEJALES for Check 3

  for (const row of rows) {
    const email = (row['e_mail'] ?? '').toLowerCase();
    const name = row['Trabajador'] ?? '';
    const fechaBaja = row['Fecha de Baja'] ?? '';

    if (isEffectiveLeaving(fechaBaja)) {
      if (email.endsWith(REQUIRED_DOMAIN)) {
        mustBeDisabled.set(email, { name, email, fechaBaja });
        allExcelEmails.add(email);
      }
    } else {
      if (email.endsWith(REQUIRED_DOMAIN)) {
        mustExist.set(email, { name, email });
        allExcelEmails.add(email);
      } else {
        noEmail.push({ name, email: email || '(empty)' });
      }
    }
  }

  // Process CONCEJALES sheet
  for (const row of concejalesRows) {
    const email = (row['email'] ?? '').toLowerCase();
    if (email && email.endsWith(REQUIRED_DOMAIN)) {
      allExcelEmails.add(email);
    }
  }

  console.log(
    chalk.gray(
      `Excel: ${mustExist.size} active employee(s) to check, ` +
      `${mustBeDisabled.size} left employee(s) to check, ` +
      `${noEmail.length} active employee(s) with no ${REQUIRED_DOMAIN} email, ` +
      `${concejalesRows.length} concejales found.\n`,
    ),
  );

  // ── 3. Fetch cloud users ───────────────────────────────────────────────────
  const spinner = ora(chalk.cyan('Fetching users from Microsoft 365…')).start();

  let cloudUsers;
  try {
    cloudUsers = await listAllUsers({ onProgress: () => { } });
  } catch (err) {
    spinner.fail(chalk.red(`\nFailed to fetch users from M365: ${err.message}`));
    process.exit(1);
  }
  spinner.succeed(chalk.cyan(`Cloud: ${cloudUsers.length} user(s) found.`));

  const cloudMap = new Map(
    cloudUsers.map((u) => [u.userPrincipalName.toLowerCase(), u]),
  );

  // ── 4. Run checks ──────────────────────────────────────────────────────────
  const missing = [];
  const notDisabled = [];
  const cloudOnly = [];

  for (const [upn, info] of mustExist) {
    if (!cloudMap.has(upn)) missing.push(info);
  }

  for (const [upn, info] of mustBeDisabled) {
    const cloudUser = cloudMap.get(upn);
    if (cloudUser && cloudUser.accountEnabled !== false) {
      notDisabled.push({ ...info, cloudStatus: 'account is ENABLED' });
    }
  }

  // Check 3 — cloud users not present in Excel (ACTIVOS or CONCEJALES)
  // Also filter out room accounts that start with 'sala-'
  for (const [upn, cloudUser] of cloudMap) {
    if (!upn.endsWith(REQUIRED_DOMAIN)) continue;
    if (upn.startsWith('sala-') || upn.startsWith('salapz')) continue; // Skip room accounts
    if (allExcelEmails.has(upn)) continue; // Skip if in ACTIVOS or CONCEJALES
    cloudOnly.push({
      name: cloudUser.displayName ?? '',
      email: upn,
      accountEnabled: cloudUser.accountEnabled,
    });
  }

  // ── 5. Print report ────────────────────────────────────────────────────────
  const totalIssues = missing.length + notDisabled.length + noEmail.length + cloudOnly.length;

  if (totalIssues === 0) {
    console.log(chalk.green('✔  Everything is in sync — no issues found!'));
    process.exit(0);
  }

  // Check 1a — missing accounts
  if (missing.length > 0) {
    console.log(chalk.bold.yellow('● Check 1a: Active employees missing from M365'));
    console.log(
      chalk.gray('  These employees have no Fecha de Baja but no matching account was found in the cloud.\n'),
    );
    for (const info of missing) {
      console.log(
        `  ${chalk.red('✖')} ${chalk.green(info.name.padEnd(40))}  ${chalk.cyan(info.email)}`,
      );
    }
    console.log('');
  }

  // Check 1b — no domain email
  if (noEmail.length > 0) {
    console.log(chalk.bold.yellow(`● Check 1b: Active employees with no ${REQUIRED_DOMAIN} email`));
    console.log(
      chalk.gray(
        '  These employees have no Fecha de Baja but their email is missing or uses a\n' +
        '  different domain, so their cloud account cannot be verified by UPN.\n',
      ),
    );
    for (const info of noEmail) {
      console.log(
        `  ${chalk.red('✖')} ${chalk.green(info.name.padEnd(40))}  ${chalk.yellow(info.email)}`,
      );
    }
    console.log('');
  }

  // Check 2 — not disabled
  if (notDisabled.length > 0) {
    console.log(chalk.bold.yellow('● Check 2: Left employees whose cloud account is not disabled'));
    console.log(
      chalk.gray(
        '  These employees have a Fecha de Baja in the file,\n' +
        '  but their cloud account is still active (or not found).\n',
      ),
    );
    for (const info of notDisabled) {
      console.log(
        `  ${chalk.red('✖')} ${chalk.green(info.name.padEnd(40))}  ` +
        `${chalk.cyan(info.email.padEnd(45))}  ` +
        `Fecha de Baja: ${chalk.yellow(info.fechaBaja.padEnd(15))}  ` +
        chalk.red(`[${info.cloudStatus}]`),
      );
    }
    console.log('');

    if (fix) {
      await fixNotDisabled(notDisabled);
      return; // fixNotDisabled calls process.exit itself
    }

    console.log(chalk.gray('  Run with --disable-left-workers to disable and unlicense these accounts.'));
    console.log('');
  }

  // Check 3 — cloud-only accounts (not in Excel)
  if (cloudOnly.length > 0) {
    console.log(chalk.bold.yellow('● Check 3: Cloud accounts not present in the Excel file'));
    console.log(
      chalk.gray(
        '  These users exist in M365 with an ' + REQUIRED_DOMAIN + ' email but do not appear\n' +
        '  in the employee Excel file (neither as active nor as left).\n',
      ),
    );
    for (const info of cloudOnly) {
      const status = info.accountEnabled === false
        ? chalk.gray('[disabled]')
        : chalk.green('[enabled]');
      console.log(
        `  ${chalk.red('✖')} ${chalk.green(info.name.padEnd(40))}  ${chalk.cyan(info.email.padEnd(45))}  ${status}`,
      );
    }
    console.log('');
  }

  console.log(chalk.red.bold(`✖  ${totalIssues} sync issue(s) found.`));
  process.exit(1);
}
