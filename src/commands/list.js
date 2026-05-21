import chalk from 'chalk';
import inquirer from 'inquirer';
import { listAllUsers, getUser, updateUser, setManager as graphSetManager, getManager } from '../graph.js';
import { editUser, editSection, FIELD_GROUPS } from './edit.js';
import { printUserCard } from './search.js';
import { stripAnsi } from '../utils/ansi.js';

/**
 * List users in the tenant with optional filters.
 *
 * @param {object} options
 * @param {boolean} options.disabled      - List disabled users instead of active ones
 * @param {boolean} options.noManager     - Only users without a manager
 * @param {boolean} options.noDepartment  - Only users without a department
 * @param {boolean} options.noJobTitle    - Only users without a job title
 * @param {boolean} options.edit          - Allow editing users from the list interactively
 */
export async function listCommand(options = {}) {
  const modeLabel = options.disabled
    ? chalk.red('disabled users')
    : chalk.green('active users');

  console.log(chalk.cyan('\nFetching users from Microsoft 365...'));
  console.log(chalk.gray(`  Mode: ${modeLabel}`));
  console.log(chalk.gray('  (Fetching manager info — may take a moment for large tenants)\n'));

  let lastPrint = 0;
  const users = await listAllUsers({
    onlyDisabled: !!options.disabled,
    checkManager: true,
    onProgress: (count) => {
      if (count - lastPrint >= 50) {
        process.stdout.write(chalk.gray(`\r  Processed ${count} users...`));
        lastPrint = count;
      }
    },
  });

  process.stdout.write('\r' + ' '.repeat(50) + '\r');

  // Client-side filters
  let filtered = users;
  if (options.noManager)    filtered = filtered.filter((u) => !u.manager);
  if (options.noDepartment) filtered = filtered.filter((u) => !u.department || u.department.trim() === '');
  if (options.noJobTitle)   filtered = filtered.filter((u) => !u.jobTitle   || u.jobTitle.trim()   === '');

  if (filtered.length === 0) {
    console.log(chalk.yellow('No users match the specified filters.'));
    return;
  }

  // Summary header
  const filterLabels = [];
  if (options.noManager)    filterLabels.push('no manager');
  if (options.noDepartment) filterLabels.push('no department');
  if (options.noJobTitle)   filterLabels.push('no job title');

  const filterStr = filterLabels.length > 0
    ? chalk.yellow(` [filters: ${filterLabels.join(', ')}]`)
    : '';

  console.log(
    `${chalk.cyan(`Found ${filtered.length} user(s)`)} ` +
    `out of ${users.length} ${options.disabled ? 'disabled' : 'active'} total` +
    filterStr + '\n'
  );

  printUserTable(filtered);

  if (!options.edit || filtered.length === 0) return;

  // Determine which section to jump to directly based on active filters
  const directSection = resolveDirectSection(options);

  await editLoop(filtered, directSection);
}

/**
 * If a single missing-field filter is active, return the section name to jump to directly.
 */
function resolveDirectSection(options) {
  const active = [
    options.noDepartment && 'Job Information',
    options.noJobTitle   && 'Job Information',
    options.noManager    && 'Manager (Org Chart)',
  ].filter(Boolean);

  // Only auto-jump when there's exactly one category of filter active
  if (active.length === 1) return active[0];
  return null;
}

/**
 * Interactive loop: pick a user from the filtered list, edit it, repeat until done.
 * Refreshes the pending-count badge next to each user as you go.
 */
async function editLoop(filtered, directSection) {
  // Track which users have been edited so we can mark them
  const edited = new Set();

  while (true) {
    const remaining = filtered.filter((u) => !edited.has(u.userPrincipalName || u.id));

    console.log(
      chalk.cyan(`\n${remaining.length} user(s) remaining`) +
      (edited.size > 0 ? chalk.green(` · ${edited.size} edited`) : '') + '\n'
    );

    const choices = filtered.map((u) => {
      const id = u.userPrincipalName || u.id;
      const done = edited.has(id);
      const missingMarks = [
        !u.jobTitle   && chalk.red('[no title]'),
        !u.department && chalk.red('[no dept]'),
        !u.manager    && chalk.red('[no manager]'),
      ].filter(Boolean).join(' ');

      return {
        name: `${done ? chalk.green('✓') : ' '} ${u.displayName} ${chalk.gray('<' + id + '>')} ${missingMarks}`,
        value: id,
      };
    });

    choices.push(new inquirer.Separator());
    choices.push({ name: chalk.bold('Done — exit'), value: '__DONE__' });

    const { selected } = await inquirer.prompt([
      {
        type: 'list',
        name: 'selected',
        message: 'Select a user to edit (or exit):',
        choices,
        pageSize: 20,
      },
    ]);

    if (selected === '__DONE__') break;

    if (directSection) {
      // Jump directly to the relevant section instead of the full editor menu
      await quickEditSection(selected, directSection);
    } else {
      await editUser(selected);
    }

    edited.add(selected);

    // Refresh the in-memory entry so missing marks update
    const idx = filtered.findIndex((u) => (u.userPrincipalName || u.id) === selected);
    if (idx !== -1) {
      const fresh = await getUser(selected).catch(() => null);
      const freshManager = await getManager(selected).catch(() => null);
      if (fresh) {
        filtered[idx] = {
          ...filtered[idx],
          jobTitle:   fresh.jobTitle,
          department: fresh.department,
          manager:    freshManager
            ? { upn: freshManager.userPrincipalName, name: freshManager.displayName }
            : null,
        };
      }
    }
  }

  console.log(chalk.green(`\nDone. ${edited.size} user(s) edited.\n`));
}

/**
 * Open only a specific section for a user and apply changes immediately.
 * Used when the context makes it obvious which fields need filling.
 */
async function quickEditSection(identifier, sectionName) {
  const user = await getUser(identifier);
  if (!user) {
    console.log(chalk.red(`User not found: ${identifier}`));
    return;
  }
  const currentManager = await getManager(identifier).catch(() => null);

  printUserCard(user, currentManager);
  console.log(chalk.cyan(`Editing section: ${chalk.bold(sectionName)}\n`));
  console.log(chalk.gray('Leave fields blank to keep current values.\n'));

  const group = FIELD_GROUPS.find((g) => g.name === sectionName);
  if (!group) {
    console.log(chalk.red(`Section not found: ${sectionName}`));
    return;
  }

  const pendingChanges = {};
  let pendingManager = null;

  await editSection(
    group,
    user,
    currentManager,
    pendingChanges,
    (m) => { pendingManager = m; },
    () => {}
  );

  const totalChanges = Object.keys(pendingChanges).length + (pendingManager !== null ? 1 : 0);

  if (totalChanges === 0) {
    console.log(chalk.yellow('No changes made.\n'));
    return;
  }

  // Apply immediately without an extra confirm step — the section prompt is confirm enough
  if (Object.keys(pendingChanges).length > 0) {
    try {
      await updateUser(identifier, pendingChanges);
      console.log(chalk.green('  Saved.'));
    } catch (err) {
      console.log(chalk.red(`  Failed to update user: ${err.message || err}`));
      return;
    }
  }

  if (pendingManager) {
    try {
      await graphSetManager(identifier, pendingManager);
      console.log(chalk.green(`  Manager set: ${pendingManager}`));
    } catch (err) {
      console.log(chalk.red(`  Failed to set manager: ${err.message || err}`));
    }
  }

  console.log();
}

/**
 * Print users as a formatted table including the manager column.
 */
function printUserTable(users) {
  const COL = {
    upn:     46,
    name:    32,
    title:   28,
    dept:    24,
    manager: 46,
  };

  const totalWidth = Object.values(COL).reduce((a, b) => a + b, 0) + Object.keys(COL).length - 1;
  const hr = chalk.gray('─'.repeat(totalWidth));

  console.log(hr);
  console.log(
    chalk.bold(pad('UPN / Email',  COL.upn))     + ' ' +
    chalk.bold(pad('Display Name', COL.name))    + ' ' +
    chalk.bold(pad('Job Title',    COL.title))   + ' ' +
    chalk.bold(pad('Department',   COL.dept))    + ' ' +
    chalk.bold(pad('Manager',      COL.manager))
  );
  console.log(hr);

  for (const u of users) {
    const titleLabel   = u.jobTitle   || chalk.red('(none)');
    const deptLabel    = u.department || chalk.red('(none)');
    const managerLabel = u.manager
      ? (u.manager.upn || u.manager.name || chalk.gray('—'))
      : chalk.red('(none)');

    console.log(
      pad(u.userPrincipalName || u.mail || u.id, COL.upn)  + ' ' +
      pad(u.displayName || '',                   COL.name)  + ' ' +
      pad(titleLabel,                            COL.title) + ' ' +
      pad(deptLabel,                             COL.dept)  + ' ' +
      pad(managerLabel,                          COL.manager)
    );
  }

  console.log(hr + '\n');
}

/** Pad or truncate a string to exactly `len` visible characters. */
function pad(str, len) {
  const plain = stripAnsi(str);
  if (plain.length >= len) return plain.slice(0, len - 1) + ' ';
  return str + ' '.repeat(len - plain.length);
}

