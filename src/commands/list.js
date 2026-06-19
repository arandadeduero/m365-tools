import chalk from 'chalk';
import inquirer from 'inquirer';
import ora from 'ora';
import { listAllUsers, getUser, updateUser, setManager as graphSetManager, getManager, fetchGroupsMap, fetchLicensesMap, getUserLicenses } from '../graph.js';
import { editUser, editSection, FIELD_GROUPS } from './edit.js';
import { printUserCard } from './search.js';
import { stripAnsi } from '../utils/ansi.js';
import { getGroupName, filterAutoGroups } from '../utils/cache.js';

/**
 * List users in the tenant with optional filters.
 *
 * @param {object} options
 * @param {boolean} options.disabled      - List disabled users instead of active ones
 * @param {boolean} options.noManager     - Only users without a manager
 * @param {boolean} options.noDepartment  - Only users without a department
 * @param {boolean} options.noJobTitle    - Only users without a job title
 * @param {boolean} options.neverSignedIn - Only users who have never signed in
 * @param {boolean} options.forMailing    - List all users with actual license, joined by ";"
 * @param {boolean} options.edit          - Allow editing users from the list interactively
 */
export async function listCommand(options = {}) {
  const modeLabel = options.disabled
    ? chalk.red('disabled users')
    : chalk.green('active users');

  const spinner = ora(chalk.cyan(`Fetching ${modeLabel} from Microsoft 365...`)).start();

  const users = await listAllUsers({
    onlyDisabled: !!options.disabled,
    checkManager: !options.forMailing,
  });

  spinner.succeed(chalk.cyan(`Fetched ${users.length} users.`));

  // If --for-mailing, we need to fetch licenses
  if (options.forMailing) {
    const licSpinner = ora(chalk.cyan('Fetching licenses for mailing list...')).start();
    for (const user of users) {
      if (!user.userPrincipalName) {
        user.hasLicense = false;
        continue;
      }
      try {
        const licenses = await getUserLicenses(user.userPrincipalName);
        user.hasLicense = licenses.length > 0;
      } catch {
        user.hasLicense = false;
      }
    }

    const emails = users
      .filter(u => u.hasLicense && u.userPrincipalName)
      .map(u => u.userPrincipalName);
    
    if (emails.length === 0) {
      licSpinner.warn(chalk.yellow('No licensed users found.'));
    } else {
      licSpinner.succeed(chalk.green('Licensed users:'));
      console.log(chalk.green(emails.join(';')));
    }
    return;
  }

  // Fetch group memberships for all users in one $batch pass.
  const groupSpinner = ora(chalk.cyan('Fetching group memberships...')).start();
  const groupsMap = await fetchGroupsMap(users);
  groupSpinner.succeed(chalk.cyan('Fetched group memberships.'));

  // Populate group name cache and attach groups array to each user,
  // filtering out auto-assigned noise groups (Todos los usuarios, etc.)
  for (const user of users) {
    const groups = filterAutoGroups(groupsMap.get(user.id) ?? []);
    user.groups = groups.map((g) => getGroupName(g.id) ?? g.displayName);
  }

  // Fetch licenses for all users
  const licSpinner = ora(chalk.cyan('Fetching licenses...')).start();
  const licensesMap = await fetchLicensesMap(users);
  licSpinner.succeed(chalk.cyan('Fetched licenses.'));

  for (const user of users) {
    user.licenses = licensesMap.get(user.id) ?? [];
  }

  // Client-side filters
  let filtered = users;
  if (options.noManager)    filtered = filtered.filter((u) => !u.manager);
  if (options.noDepartment) filtered = filtered.filter((u) => !u.department || u.department.trim() === '');
  if (options.noJobTitle)   filtered = filtered.filter((u) => !u.jobTitle   || u.jobTitle.trim()   === '');
  if (options.neverSignedIn) filtered = filtered.filter((u) => !u.signInActivity?.lastSignInDateTime);

  if (filtered.length === 0) {
    console.log(chalk.yellow('No users match the specified filters.'));
    return;
  }

  // Summary header
  const filterLabels = [];
  if (options.noManager)    filterLabels.push('no manager');
  if (options.noDepartment) filterLabels.push('no department');
  if (options.noJobTitle)   filterLabels.push('no job title');
  if (options.neverSignedIn) filterLabels.push('never signed in');

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
        type: 'select',
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
 * Print users as a formatted table including employee ID, manager, groups, and licenses.
 */
function printUserTable(users) {
  const COL = {
    empId:   10,
    upn:     38,
    name:    28,
    title:   24,
    dept:    20,
    manager: 36,
    groups:  40,
    licenses: 30, // Added
    lastSignIn: 20,
    lastPasswordChange: 20,
  };

  const totalWidth = Object.values(COL).reduce((a, b) => a + b, 0) + Object.keys(COL).length - 1;
  const hr = chalk.gray('─'.repeat(totalWidth));

  console.log(hr);
  console.log(
    chalk.bold(pad('Emp ID',       COL.empId))   + ' ' +
    chalk.bold(pad('UPN / Email',  COL.upn))     + ' ' +
    chalk.bold(pad('Display Name', COL.name))    + ' ' +
    chalk.bold(pad('Job Title',    COL.title))   + ' ' +
    chalk.bold(pad('Department',   COL.dept))    + ' ' +
    chalk.bold(pad('Manager',      COL.manager)) + ' ' +
    chalk.bold(pad('Groups',       COL.groups))  + ' ' +
    chalk.bold(pad('Licenses',     COL.licenses)) + ' ' +
    chalk.bold(pad('Last Sign-In', COL.lastSignIn)) + ' ' +
    chalk.bold(pad('Last Pass Change', COL.lastPasswordChange))
  );
  console.log(hr);

  for (const u of users) {
    const empIdLabel   = u.employeeId  || chalk.gray('—');
    const titleLabel   = u.jobTitle    || chalk.red('(none)');
    const deptLabel    = u.department  || chalk.red('(none)');
    const managerLabel = u.manager
      ? (u.manager.upn || u.manager.name || chalk.gray('—'))
      : chalk.red('(none)');
    const groupsLabel  = u.groups && u.groups.length > 0
      ? u.groups.join(', ')
      : chalk.gray('—');
    const licensesLabel = u.licenses && u.licenses.length > 0
      ? u.licenses.map(l => l.skuPartNumber).join(', ')
      : chalk.gray('—');
    const lastSignInLabel = u.signInActivity?.lastSignInDateTime
      ? u.signInActivity.lastSignInDateTime.split('T')[0]
      : chalk.gray('—');
    const lastPasswordChangeLabel = u.lastPasswordChangeDateTime
      ? u.lastPasswordChangeDateTime.split('T')[0]
      : chalk.gray('—');

    console.log(
      pad(empIdLabel,                            COL.empId)   + ' ' +
      pad(u.userPrincipalName || u.mail || u.id, COL.upn)     + ' ' +
      pad(u.displayName || '',                   COL.name)    + ' ' +
      pad(titleLabel,                            COL.title)   + ' ' +
      pad(deptLabel,                             COL.dept)    + ' ' +
      pad(managerLabel,                          COL.manager) + ' ' +
      pad(groupsLabel,                           COL.groups)  + ' ' +
      pad(licensesLabel,                         COL.licenses) + ' ' +
      pad(lastSignInLabel,                       COL.lastSignIn) + ' ' +
      pad(lastPasswordChangeLabel,               COL.lastPasswordChange)
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

