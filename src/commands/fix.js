import chalk from 'chalk';
import inquirer from 'inquirer';
import ora from 'ora';
import { listAllUsers, updateUser, fetchManagerMap, setManager, getUser } from '../graph.js';
import { getDomain } from '../auth.js';
import { normalizeUpn } from '../utils/upn.js';
import { stripAnsi } from '../utils/ansi.js';

/**
 * Convert a string to Sentence case:
 * First character uppercase, rest lowercase.
 * Preserves empty / null values.
 */
function toSentenceCase(str) {
  if (!str || str.trim() === '') return str;
  const s = str.trim();
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

/**
 * Fix job titles: normalise all active users' jobTitle to Sentence case.
 * Shows a preview table, lets the user pick which ones to apply, then confirms twice.
 */
export async function fixJobTitles() {
  const spinner = ora(chalk.cyan('Fetching all active users with a job title...')).start();

  const allUsers = await listAllUsers({
    onlyDisabled: false,
    checkManager: false,
  });

  // Keep only users whose jobTitle would actually change
  const candidates = allUsers
    .filter((u) => u.jobTitle && u.jobTitle.trim() !== '')
    .map((u) => ({
      id:          u.id,
      upn:         u.userPrincipalName || u.id,
      displayName: u.displayName || '',
      original:    u.jobTitle,
      fixed:       toSentenceCase(u.jobTitle),
    }))
    .filter((u) => u.original !== u.fixed);

  if (candidates.length === 0) {
    spinner.succeed(chalk.green('All job titles are already in Sentence case. Nothing to do.'));
    return;
  }
  spinner.succeed(chalk.cyan(`Found ${candidates.length} candidates.`));

  // Step 1 — let user pick which ones to apply (all selected by default)
  const { selected } = await inquirer.prompt([
    {
      type: 'checkbox',
      name: 'selected',
      message: 'Select users to apply the fix to (Space to toggle, A to select all):',
      choices: candidates.map((u) => ({
        name: `${u.displayName} ${chalk.gray('<' + u.upn + '>')}  ${chalk.red(u.original)} → ${chalk.green(u.fixed)}`,
        value: u.id,
        checked: true, // all selected by default
      })),
      pageSize: 20,
    },
  ]);

  if (selected.length === 0) {
    console.log(chalk.yellow('\nNo users selected. Nothing to do.\n'));
    return;
  }

  const toApply = candidates.filter((u) => selected.includes(u.id));

  // Double-check #1
  console.log(chalk.yellow(`\nYou are about to update ${toApply.length} user(s):`));
  for (const u of toApply) {
    console.log(`  ${chalk.bold(u.displayName)} — ${chalk.red(u.original)} → ${chalk.green(u.fixed)}`);
  }

  const { confirm1 } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm1',
      message: `Apply Sentence case fix to ${toApply.length} user(s)?`,
      default: false,
    },
  ]);

  if (!confirm1) {
    console.log(chalk.gray('\nCancelled.\n'));
    return;
  }

  // Double-check #2
  const { confirm2 } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm2',
      message: chalk.yellow('Are you sure? This will overwrite job titles in Microsoft 365.'),
      default: false,
    },
  ]);

  if (!confirm2) {
    console.log(chalk.gray('\nCancelled.\n'));
    return;
  }

  // Apply in batches of 10
  const applySpinner = ora(chalk.cyan('Applying changes...')).start();

  const BATCH = 10;
  const results = { ok: [], failed: [] };

  for (let i = 0; i < toApply.length; i += BATCH) {
    const slice = toApply.slice(i, i + BATCH);
    const settled = await Promise.allSettled(
      slice.map((u) => updateUser(u.id, { jobTitle: u.fixed }))
    );
    for (const [idx, r] of settled.entries()) {
      const u = slice[idx];
      if (r.status === 'fulfilled') {
        results.ok.push(u);
      } else {
        results.failed.push({ ...u, error: r.reason?.message || String(r.reason) });
      }
    }
    applySpinner.text = chalk.cyan(`Applying changes... (${Math.min(i + BATCH, toApply.length)}/${toApply.length})`);
  }
  applySpinner.succeed(chalk.cyan('Changes applied.'));

  // Summary
  console.log(chalk.cyan('\n─────────────────────────────────────────'));
  console.log(chalk.bold('Summary:'));
  console.log(`  ${chalk.green(`Updated: ${results.ok.length}`)}`);
  console.log(`  ${chalk.red(`Failed:  ${results.failed.length}`)}`);
  console.log(chalk.cyan('─────────────────────────────────────────\n'));

  if (results.failed.length > 0) {
    console.log(chalk.red('Failed users:'));
    for (const f of results.failed) {
      console.log(chalk.red(`  - ${f.displayName} <${f.upn}>: ${f.error}`));
    }
    console.log();
  }
}

/**
 * Print a preview table: UPN | Display Name | Original title | Fixed title
 */
function printPreviewTable(candidates) {
  const COL = { upn: 42, name: 28, original: 32, fixed: 32 };
  const totalWidth = Object.values(COL).reduce((a, b) => a + b, 0) + Object.keys(COL).length - 1;
  const hr = chalk.gray('─'.repeat(totalWidth));

  console.log(hr);
  console.log(
    chalk.bold(pad('UPN / Email',    COL.upn))      + ' ' +
    chalk.bold(pad('Display Name',   COL.name))     + ' ' +
    chalk.bold(pad('Current Title',  COL.original)) + ' ' +
    chalk.bold(pad('→ Fixed Title',  COL.fixed))
  );
  console.log(hr);

  for (const u of candidates) {
    console.log(
      pad(u.upn,          COL.upn)      + ' ' +
      pad(u.displayName,  COL.name)     + ' ' +
      pad(chalk.red(u.original),  COL.original) + ' ' +
      pad(chalk.green(u.fixed),   COL.fixed)
    );
  }

  console.log(hr + '\n');
}

function pad(str, len) {
  const plain = stripAnsi(str);
  if (plain.length >= len) return plain.slice(0, len - 1) + ' ';
  return str + ' '.repeat(len - plain.length);
}

// ─────────────────────────────────────────────────────────────────────────────
// ORG CHART CHECK
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Check that the org chart is valid:
 *   1. Every active user has a manager.
 *   2. Following the manager chain from any user always reaches a single root.
 *   3. There are no cycles.
 *
 * If problems are found, lets the user fix them interactively.
 */
export async function fixOrgChart() {
  // ── Step 1: fetch all active users ────────────────────────────────────────
  const userSpinner = ora(chalk.cyan('Fetching all active users...')).start();
  const users = await listAllUsers({
    onlyDisabled: false,
    checkManager: false,
  });
  userSpinner.succeed(chalk.cyan(`Loaded ${users.length} active users.`));

  // Build lookup maps
  const byId  = new Map(users.map((u) => [u.id, u]));
  const byUpn = new Map(users.map((u) => [u.userPrincipalName, u]));

  // ── Step 2: fetch manager for every user via $batch ───────────────────────
  const managerSpinner = ora(chalk.cyan('Fetching manager relationships ($batch)...')).start();
  const managerMap = await fetchManagerMap(users);
  managerSpinner.succeed(chalk.cyan('Manager relationships loaded.'));

  // ── Step 3: ask for the expected root UPN ─────────────────────────────────
  const domain = getDomain();
  const { rootInput } = await inquirer.prompt([{
    type:    'input',
    name:    'rootInput',
    message: `Enter the UPN of the expected org root (e.g. alinaje${domain ? '@' + domain : ''})${domain ? chalk.gray(' [@' + domain + ' added if missing]') : ''}:`,
    filter:  (v) => normalizeUpn(v, domain),
    validate: (v) => v && v.includes('@') ? true : 'Please enter a valid UPN.',
  }]);

  // Resolve root user
  const rootUser = byUpn.get(rootInput) || await getUser(rootInput).catch(() => null);
  if (!rootUser) {
    console.log(chalk.red(`\nRoot user not found: ${rootInput}`));
    return;
  }
  const rootId = rootUser.id;
  console.log(chalk.green(`\n  Root: ${rootUser.displayName} <${rootUser.userPrincipalName}>\n`));

  // ── Step 4: analyse the tree ──────────────────────────────────────────────
  const issues = analyseOrgTree(users, managerMap, rootId, byId);

  if (issues.noManager.length === 0 && issues.wrongRoot.length === 0 && issues.cycles.length === 0) {
    console.log(chalk.green('✓ Org chart is valid. All users connect to the root with no cycles.\n'));
    return;
  }

  printIssuesSummary(issues, byId);

  // ── Step 5: let user fix the problems ─────────────────────────────────────
  const { action } = await inquirer.prompt([{
    type:    'list',
    name:    'action',
    message: 'What do you want to do?',
    choices: [
      { name: 'Fix interactively — assign a manager to each affected user', value: 'fix' },
      { name: 'Exit without changes',                                        value: 'exit' },
    ],
  }]);

  if (action === 'exit') return;

  // Collect all problem users (deduplicated)
  const problemIds = new Set([
    ...issues.noManager.map((u) => u.id),
    ...issues.wrongRoot.map((u) => u.id),
    ...issues.cycles.flat().map((u) => u.id),
  ]);

  const problemUsers = [...problemIds]
    .map((id) => byId.get(id))
    .filter(Boolean)
    .sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));

  await orgFixLoop(problemUsers, managerMap, byId, domain);
}

// ─────────────────────────────────────────────────────────────────────────────
// Tree analysis
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Analyse the manager tree and return categorised issues.
 *
 * @returns {{ noManager: User[], wrongRoot: User[], cycles: User[][] }}
 */
function analyseOrgTree(users, managerMap, rootId, byId) {
  const noManager  = [];
  const wrongRoot  = [];
  const cycles     = [];
  const cycleIds   = new Set();

  for (const user of users) {
    if (user.id === rootId) continue; // root has no manager by design

    const managerId = managerMap.get(user.id);

    // 1. No manager at all
    if (!managerId) {
      noManager.push(user);
      continue;
    }

    // 2. Check chain leads to root (max 50 hops to avoid infinite loop)
    const visited  = new Set();
    let   current  = user.id;
    let   reachesRoot = false;
    let   hasCycle    = false;

    for (let hop = 0; hop < 50; hop++) {
      const nextId = managerMap.get(current);
      if (!nextId) break;                  // chain ends without reaching root
      if (nextId === rootId) { reachesRoot = true; break; }
      if (visited.has(nextId)) { hasCycle = true; break; }
      visited.add(nextId);
      current = nextId;
    }

    if (hasCycle && !cycleIds.has(user.id)) {
      // Collect the cycle members
      const cycleMembers = [user];
      let c = managerMap.get(user.id);
      const seen = new Set([user.id]);
      while (c && !seen.has(c) && byId.has(c)) {
        seen.add(c);
        cycleMembers.push(byId.get(c));
        c = managerMap.get(c);
      }
      cycles.push(cycleMembers);
      cycleMembers.forEach((u) => cycleIds.add(u.id));
      continue;
    }

    if (!reachesRoot) {
      wrongRoot.push(user);
    }
  }

  return { noManager, wrongRoot, cycles };
}

// ─────────────────────────────────────────────────────────────────────────────
// Reporting
// ─────────────────────────────────────────────────────────────────────────────

function printIssuesSummary(issues, byId) {
  const COL = { name: 30, upn: 44, issue: 30 };
  const hr  = chalk.gray('─'.repeat(COL.name + COL.upn + COL.issue + 2));

  console.log(hr);
  console.log(
    chalk.bold(pad('Display Name', COL.name)) + ' ' +
    chalk.bold(pad('UPN',          COL.upn))  + ' ' +
    chalk.bold('Issue')
  );
  console.log(hr);

  const printRow = (u, issue) =>
    console.log(
      pad(u.displayName || '', COL.name) + ' ' +
      pad(u.userPrincipalName || u.id, COL.upn) + ' ' +
      issue
    );

  for (const u of issues.noManager) {
    printRow(u, chalk.red('No manager'));
  }
  for (const u of issues.wrongRoot) {
    printRow(u, chalk.yellow('Chain does not reach root'));
  }
  for (const cycle of issues.cycles) {
    for (const u of cycle) {
      printRow(u, chalk.magenta('Cycle detected'));
    }
  }

  console.log(hr);

  const total = issues.noManager.length + issues.wrongRoot.length +
    issues.cycles.reduce((s, c) => s + c.length, 0);

  console.log(
    `\n  ${chalk.red(`No manager: ${issues.noManager.length}`)}` +
    `  ${chalk.yellow(`Wrong root: ${issues.wrongRoot.length}`)}` +
    `  ${chalk.magenta(`Cycles: ${issues.cycles.length}`)}` +
    `  ${chalk.bold(`Total affected: ${total}`)}\n`
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Interactive fix loop
// ─────────────────────────────────────────────────────────────────────────────

async function orgFixLoop(problemUsers, managerMap, byId, domain) {
  const fixed = new Set();

  while (true) {
    const remaining = problemUsers.filter((u) => !fixed.has(u.id));
    if (remaining.length === 0) {
      console.log(chalk.green('\n✓ All issues resolved.\n'));
      break;
    }

    console.log(chalk.cyan(`\n${remaining.length} user(s) still need a manager.\n`));

    const choices = problemUsers.map((u) => ({
      name: `${fixed.has(u.id) ? chalk.green('✓') : ' '} ${pad(u.displayName || '', 28)} ${chalk.gray(u.userPrincipalName || u.id)}`,
      value: u.id,
    }));
    choices.push(new inquirer.Separator());
    choices.push({ name: chalk.bold('Done — exit'), value: '__DONE__' });

    const { selectedId } = await inquirer.prompt([{
      type:     'list',
      name:     'selectedId',
      message:  'Select a user to assign a manager to:',
      choices,
      pageSize: 20,
    }]);

    if (selectedId === '__DONE__') break;

    const user = byId.get(selectedId);
    if (!user) continue;

    console.log(
      chalk.cyan(`\nAssigning manager for: `) +
      chalk.bold(user.displayName) + chalk.gray(` <${user.userPrincipalName}>\n`)
    );

    const { managerInput } = await inquirer.prompt([{
      type:    'input',
      name:    'managerInput',
      message: `Manager UPN${domain ? chalk.gray(' [@' + domain + ' added if missing]') : ''}:`,
      filter:  (v) => normalizeUpn(v, domain),
      validate: async (v) => {
        if (!v || !v.includes('@')) return 'Enter a valid UPN.';
        if (v === user.userPrincipalName) return 'A user cannot be their own manager.';
        process.stdout.write(chalk.gray(' Verifying...'));
        const found = await getUser(v).catch(() => null);
        return found ? true : `User not found: ${v}`;
      },
    }]);

    try {
      await setManager(user.id, managerInput);
      console.log(chalk.green(`  Manager set: ${managerInput}\n`));
      fixed.add(user.id);
    } catch (err) {
      console.log(chalk.red(`  Failed: ${err.message || err}\n`));
    }
  }

  if (fixed.size > 0) {
    console.log(chalk.green(`\nFixed ${fixed.size} user(s). Run the check again to verify.\n`));
  }
}
