import chalk from 'chalk';
import inquirer from 'inquirer';
import { listAllUsers, getMailboxTzSettings, setMailboxTzSettings } from '../graph.js';
import { getAccessToken } from '../auth.js';
import { stripAnsi } from '../utils/ansi.js';

const TARGET_TZ = 'Romance Standard Time'; // Europe/Madrid (UTC+1 / UTC+2 DST)

// ---------------------------------------------------------------------------
// Public command
// ---------------------------------------------------------------------------

/**
 * Set the mailbox timezone to Romance Standard Time (Europe/Madrid) for all
 * active users whose current mailbox timezone differs from the target.
 *
 * Fixes both mailboxSettings.timeZone (calendar display) AND
 * mailboxSettings.workingHours.timeZone (Teams free/busy, availability) in a
 * single PATCH per user. The working hours schedule (days, start/end times)
 * is preserved exactly as-is for each user.
 *
 * Requires the authenticated account to have the Exchange Administrator or
 * Global Administrator role in Azure AD, so that delegated
 * MailboxSettings.ReadWrite grants access to other users' mailboxes.
 *
 * @param {object}  opts
 * @param {boolean} opts.dryRun - Preview changes without applying them
 * @param {boolean} opts.debug  - Print raw API response for each user
 */
export async function fixTimezone({ dryRun = false, debug = false } = {}) {

  // ── Debug: show token scopes and test one raw mailboxSettings call ─────────
  if (debug) {
    const token = await getAccessToken();
    try {
      const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString('utf8'));
      console.log(chalk.cyan('\n[debug] Token info:'));
      console.log(chalk.gray('  upn:    ') + (payload.upn || payload.unique_name || '(none)'));
      console.log(chalk.gray('  scopes: ') + (payload.scp || payload.scope || '(none)'));
      console.log(chalk.gray('  roles:  ') + (payload.roles?.join(', ') || '(none)'));
      console.log(chalk.gray('  exp:    ') + new Date((payload.exp ?? 0) * 1000).toISOString() + '\n');
    } catch {
      console.log(chalk.yellow('[debug] Could not decode token payload.\n'));
    }

    console.log(chalk.cyan('[debug] Testing raw mailboxSettings fetch for /me ...'));
    try {
      const token2 = await getAccessToken();
      const res = await fetch('https://graph.microsoft.com/v1.0/me/mailboxSettings', {
        headers: { Authorization: `Bearer ${token2}` },
      });
      const body = await res.text();
      console.log(chalk.gray(`  HTTP ${res.status}`));
      console.log(chalk.gray('  body: ') + body.slice(0, 500) + '\n');
    } catch (e) {
      console.log(chalk.red('  fetch error: ' + e.message + '\n'));
    }
  }

  console.log(chalk.cyan('\nFetching all active users...\n'));

  let lastPrint = 0;
  const allUsers = await listAllUsers({
    onlyDisabled: false,
    checkManager: false,
    onProgress: (count) => {
      if (count - lastPrint >= 100) {
        process.stdout.write(chalk.gray(`\r  Fetched ${count} users...`));
        lastPrint = count;
      }
    },
  });

  process.stdout.write('\r' + ' '.repeat(50) + '\r');
  console.log(chalk.gray(`  ${allUsers.length} active users loaded.\n`));

  // ── Read current mailbox timezone settings for every user ─────────────────
  // Done sequentially to stay well within Exchange Online throttle limits.
  console.log(chalk.cyan('Reading mailbox timezone settings...\n'));

  const alreadyCorrect = [];
  const needsUpdate    = [];
  const noMailbox      = [];
  const accessDenied   = [];

  for (let i = 0; i < allUsers.length; i++) {
    const u = allUsers[i];
    if (!debug) {
      process.stdout.write(
        chalk.gray(`\r  [${i + 1}/${allUsers.length}] ${pad(u.userPrincipalName || u.id, 50)}`)
      );
    }

    const settings = await getMailboxTzSettings(u.id);

    if (debug) {
      console.log(chalk.gray(`[debug] ${u.userPrincipalName || u.id}`));
      if (settings?._error) {
        console.log(chalk.red('  error: ') + JSON.stringify(settings._error, null, 2)
          .split('\n').map((l) => '  ' + l).join('\n'));
      } else {
        console.log(chalk.gray('  raw: ') + JSON.stringify(settings?._raw ?? null, null, 2)
          .split('\n').map((l) => '  ' + l).join('\n'));
      }
    }

    if (settings === null) {
      // Should not happen with new code, but guard anyway
      noMailbox.push({ ...u, settings: { _reason: 'no-mailbox' } });
    } else if (settings._error) {
      if (settings._reason === 'access-denied') {
        accessDenied.push({ ...u, settings });
      } else {
        noMailbox.push({ ...u, settings });
      }
    } else {
      const tzOk   = settings.timeZone === TARGET_TZ;
      const whTzOk = settings.workingHoursTimeZone === null ||
                     settings.workingHoursTimeZone === TARGET_TZ;

      if (tzOk && whTzOk) {
        alreadyCorrect.push({ ...u, settings });
      } else {
        needsUpdate.push({ ...u, settings });
      }
    }

    // Small pause to respect Exchange Online request-rate limits
    await sleep(100);
  }

  if (!debug) process.stdout.write('\r' + ' '.repeat(60) + '\r');

  // ── Summary ───────────────────────────────────────────────────────────────
  console.log(chalk.cyan('─────────────────────────────────────────'));
  console.log(chalk.bold('Mailbox timezone audit:'));
  console.log(`  ${chalk.green(`Already correct (${TARGET_TZ}): ${alreadyCorrect.length}`)}`);
  console.log(`  ${chalk.yellow(`Need update:                    ${needsUpdate.length}`)}`);
  console.log(`  ${chalk.gray(`No mailbox / no license:        ${noMailbox.length}`)}`);
  if (accessDenied.length > 0) {
    console.log(`  ${chalk.red(`Access denied (needs role):     ${accessDenied.length}`)}`);
  }
  console.log(chalk.cyan('─────────────────────────────────────────\n'));

  // ── Access denied warning ─────────────────────────────────────────────────
  if (accessDenied.length > 0) {
    console.log(chalk.red('  ┌─ ACCESS DENIED for ' + accessDenied.length + ' user(s) ──────────────────────────────────────┐'));
    console.log(chalk.red('  │'));
    console.log(chalk.red('  │  The authenticated account lacks permission to read/write other'));
    console.log(chalk.red('  │  users\' mailbox settings. This is an Exchange-level restriction.'));
    console.log(chalk.red('  │'));
    console.log(chalk.red('  │  Fix: assign the ') + chalk.bold.red('Exchange Administrator') + chalk.red(' role to:'));
    console.log(chalk.red('  │    ') + chalk.yellow(allUsers.find(u => u.userPrincipalName)?.userPrincipalName?.replace(/.*@/, '<your-admin>@') || '<your-admin-account>'));
    console.log(chalk.red('  │'));
    console.log(chalk.red('  │  Azure Portal → Azure Active Directory → Roles and administrators'));
    console.log(chalk.red('  │  → Exchange Administrator → + Add assignments → pick your account'));
    console.log(chalk.red('  │'));
    console.log(chalk.red('  │  Then run: ') + chalk.bold('m365-users logout && m365-users login'));
    console.log(chalk.red('  └──────────────────────────────────────────────────────────────────┘\n'));
  }

  if (needsUpdate.length === 0 && accessDenied.length === 0) {
    console.log(chalk.green(`All reachable mailboxes are already set to ${TARGET_TZ}. Nothing to do.\n`));
    return;
  }

  if (needsUpdate.length === 0) {
    // Only access-denied, nothing actionable from code side
    return;
  }

  // ── Preview table ─────────────────────────────────────────────────────────
  printPreviewTable(needsUpdate);

  if (dryRun) {
    console.log(chalk.yellow(`[dry-run] ${needsUpdate.length} user(s) would be updated. No changes were made.\n`));
    return;
  }

  // ── Confirmation #1 ───────────────────────────────────────────────────────
  const { confirm1 } = await inquirer.prompt([
    {
      type:    'confirm',
      name:    'confirm1',
      message: `Set mailbox timezone to "${TARGET_TZ}" for ${needsUpdate.length} user(s)?`,
      default: false,
    },
  ]);

  if (!confirm1) {
    console.log(chalk.gray('\nCancelled.\n'));
    return;
  }

  // ── Confirmation #2 ───────────────────────────────────────────────────────
  const { confirm2 } = await inquirer.prompt([
    {
      type:    'confirm',
      name:    'confirm2',
      message: chalk.yellow('Are you sure? This will overwrite mailbox timezone settings in Microsoft 365.'),
      default: false,
    },
  ]);

  if (!confirm2) {
    console.log(chalk.gray('\nCancelled.\n'));
    return;
  }

  // ── Apply changes sequentially ────────────────────────────────────────────
  console.log(chalk.cyan('\nApplying changes...\n'));

  const results = { ok: [], failed: [] };

  for (const u of needsUpdate) {
    const { settings } = u;
    try {
      await setMailboxTzSettings(u.id, TARGET_TZ, settings.workingHours);
      results.ok.push(u);

      const changed = [];
      if (settings.timeZone !== TARGET_TZ)
        changed.push(`calendar: ${chalk.red(settings.timeZone || '(unset)')}`);
      if (settings.workingHoursTimeZone !== null && settings.workingHoursTimeZone !== TARGET_TZ)
        changed.push(`working hours: ${chalk.red(settings.workingHoursTimeZone)}`);

      console.log(
        chalk.green('  [OK]    ') +
        chalk.bold(pad(u.displayName || '', 30)) + '  ' +
        changed.join(', ') + chalk.green(' → ' + TARGET_TZ)
      );
    } catch (err) {
      results.failed.push({ ...u, error: err.message || String(err) });
      console.log(
        chalk.red('  [FAIL]  ') +
        chalk.bold(pad(u.displayName || '', 30)) + '  ' +
        chalk.red(err.message || err)
      );
    }

    await sleep(100);
  }

  // ── Final summary ─────────────────────────────────────────────────────────
  console.log(chalk.cyan('\n─────────────────────────────────────────'));
  console.log(chalk.bold('Summary:'));
  console.log(`  ${chalk.green(`Updated:           ${results.ok.length}`)}`);
  console.log(`  ${chalk.gray(`Already correct:   ${alreadyCorrect.length}`)}`);
  console.log(`  ${chalk.gray(`No mailbox:        ${noMailbox.length}`)}`);
  if (accessDenied.length > 0)
    console.log(`  ${chalk.red(`Access denied:     ${accessDenied.length}`)}`);
  console.log(`  ${chalk.red(`Failed:            ${results.failed.length}`)}`);
  console.log(chalk.cyan('─────────────────────────────────────────\n'));

  if (results.failed.length > 0) {
    console.log(chalk.red('Failed users:'));
    for (const f of results.failed) {
      console.log(chalk.red(`  - ${f.displayName} <${f.userPrincipalName}>: ${f.error}`));
    }
    console.log();
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function printPreviewTable(candidates) {
  const COL = { upn: 42, name: 26, tz: 26, whTz: 26 };
  const totalWidth = Object.values(COL).reduce((a, b) => a + b, 0) + Object.keys(COL).length - 1;
  const hr = chalk.gray('─'.repeat(totalWidth));

  console.log(hr);
  console.log(
    chalk.bold(pad('UPN / Email',      COL.upn))  + ' ' +
    chalk.bold(pad('Display Name',     COL.name)) + ' ' +
    chalk.bold(pad('Calendar TZ',      COL.tz))   + ' ' +
    chalk.bold(pad('Working Hours TZ', COL.whTz))
  );
  console.log(hr);

  for (const u of candidates) {
    const { settings } = u;
    const tzVal = settings.timeZone === TARGET_TZ
      ? chalk.green(TARGET_TZ)
      : chalk.red(settings.timeZone || '(unset)');
    const whVal = settings.workingHoursTimeZone === null
      ? chalk.gray('(none)')
      : settings.workingHoursTimeZone === TARGET_TZ
        ? chalk.green(TARGET_TZ)
        : chalk.red(settings.workingHoursTimeZone);

    console.log(
      pad(u.userPrincipalName || u.id, COL.upn)  + ' ' +
      pad(u.displayName || '',          COL.name) + ' ' +
      pad(tzVal,                        COL.tz)   + ' ' +
      pad(whVal,                        COL.whTz)
    );
  }

  console.log(hr + '\n');
}

function pad(str, len) {
  const plain = stripAnsi(str);
  if (plain.length >= len) return plain.slice(0, len - 1) + ' ';
  return str + ' '.repeat(len - plain.length);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
