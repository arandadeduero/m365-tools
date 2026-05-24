import chalk from 'chalk';
import inquirer from 'inquirer';
import {
  listAllJobTitles,
  listAllDepartments,
  getUsersByJobTitle,
  getUsersByDepartment,
  getUser,
  setManager,
} from '../graph.js';
import { getDomain } from '../auth.js';
import { normalizeUpn } from '../utils/upn.js';

/**
 * Interactive command: choose filter mode (job title or department),
 * pick a value from the real list, enter a manager UPN, and bulk-assign.
 */
export async function assignManagerByJobTitle() {
  // Step 1 — choose filter mode
  const { mode } = await inquirer.prompt([
    {
      type: 'select',
      name: 'mode',
      message: 'Assign manager by:',
      choices: [
        { name: 'Job Title   — e.g. "AGENTE DE POLICÍA"', value: 'jobTitle' },
        { name: 'Department  — e.g. "URBANISMO"',         value: 'department' },
      ],
    },
  ]);

  // Step 2 — fetch available values and let user pick
  let users;
  let filterLabel;

  if (mode === 'jobTitle') {
    console.log(chalk.cyan('\nFetching all job titles from active users...'));
    console.log(chalk.gray('  (This may take a moment)\n'));

    const titles = await listAllJobTitles();
    if (titles.length === 0) {
      console.log(chalk.yellow('No job titles found in the tenant.'));
      return;
    }
    console.log(chalk.gray(`  Found ${titles.length} distinct job title(s).\n`));

    const { jobTitle } = await inquirer.prompt([
      {
        type: 'select',
        name: 'jobTitle',
        message: 'Select job title:',
        choices: titles,
        pageSize: 20,
      },
    ]);

    filterLabel = `job title "${jobTitle}"`;
    console.log(chalk.cyan(`\nFetching users with ${filterLabel}...`));
    users = await getUsersByJobTitle(jobTitle);

  } else {
    console.log(chalk.cyan('\nFetching all departments from active users...'));
    console.log(chalk.gray('  (This may take a moment)\n'));

    const depts = await listAllDepartments();
    if (depts.length === 0) {
      console.log(chalk.yellow('No departments found in the tenant.'));
      return;
    }
    console.log(chalk.gray(`  Found ${depts.length} distinct department(s).\n`));

    const { department } = await inquirer.prompt([
      {
        type: 'select',
        name: 'department',
        message: 'Select department:',
        choices: depts,
        pageSize: 20,
      },
    ]);

    filterLabel = `department "${department}"`;
    console.log(chalk.cyan(`\nFetching users in ${filterLabel}...`));
    users = await getUsersByDepartment(department);
  }

  // Step 3 — show matched users
  if (users.length === 0) {
    console.log(chalk.yellow(`No active users found for ${filterLabel}.`));
    return;
  }

  console.log(chalk.cyan(`\nFound ${users.length} user(s) for ${filterLabel}:\n`));

  for (const u of users) {
    const meta = [u.jobTitle, u.department].filter(Boolean).join(' — ');
    console.log(
      `  ${chalk.bold(u.displayName)} — ${chalk.gray(u.userPrincipalName)}` +
      (meta ? chalk.gray(`  [${meta}]`) : '')
    );
  }

  // Step 4 — ask for manager UPN
  const { managerUpn } = await inquirer.prompt([
    {
      type: 'input',
      name: 'managerUpn',
      message: `\nEnter the manager UPN (email)${getDomain() ? ` [@${getDomain()} added automatically if missing]` : ''}:`,
      validate: (v) => {
        if (!v || !v.trim()) return 'Manager UPN cannot be empty.';
        return true;
      },
      filter: (v) => normalizeUpn(v, getDomain()),
    },
  ]);

  // Step 5 — resolve manager
  console.log(chalk.gray(`\nLooking up manager "${managerUpn.trim()}"...`));
  const manager = await getUser(managerUpn.trim());

  if (!manager) {
    console.log(chalk.red(`Manager not found: ${managerUpn.trim()}`));
    return;
  }

  console.log(chalk.green(`  Manager found: ${manager.displayName} <${manager.userPrincipalName}>\n`));

  // Step 6 — confirm
  const { confirm } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm',
      message: `Assign "${manager.displayName}" as manager to all ${users.length} user(s) with ${filterLabel}?`,
      default: false,
    },
  ]);

  if (!confirm) {
    console.log(chalk.gray('\nCancelled. No changes made.\n'));
    return;
  }

  // Step 7 — bulk assign in parallel batches of 10
  console.log(chalk.cyan('\nAssigning manager...\n'));

  const BATCH = 10;
  const results = { ok: [], failed: [] };

  for (let i = 0; i < users.length; i += BATCH) {
    const slice = users.slice(i, i + BATCH);
    const settled = await Promise.allSettled(
      slice.map((u) => setManager(u.id, manager.userPrincipalName, manager.id))
    );
    settled.forEach((r, idx) => {
      const u = slice[idx];
      if (r.status === 'fulfilled') {
        results.ok.push(u.userPrincipalName);
        console.log(chalk.green(`  [OK]     ${u.displayName} <${u.userPrincipalName}>`));
      } else {
        results.failed.push({ upn: u.userPrincipalName, error: r.reason?.message || String(r.reason) });
        console.log(chalk.red(`  [FAILED] ${u.displayName} <${u.userPrincipalName}>: ${r.reason?.message || r.reason}`));
      }
    });
  }

  // Summary
  console.log(chalk.cyan('\n─────────────────────────────────────────'));
  console.log(chalk.bold('Summary:'));
  console.log(`  ${chalk.green(`Updated: ${results.ok.length}`)}`);
  console.log(`  ${chalk.red(`Failed:  ${results.failed.length}`)}`);
  console.log(chalk.cyan('─────────────────────────────────────────\n'));

  if (results.failed.length > 0) {
    console.log(chalk.red('Failed users:'));
    for (const f of results.failed) {
      console.log(chalk.red(`  - ${f.upn}: ${f.error}`));
    }
    console.log();
  }
}
