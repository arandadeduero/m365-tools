import chalk from 'chalk';
import inquirer from 'inquirer';
import ora from 'ora';
import { searchUsers, getUser, getManager } from '../graph.js';
import { editUser } from './edit.js';

/**
 * Search for users by name or email and optionally open the editor.
 * @param {string} query - Search query (name or email)
 * @param {object} options - { edit: boolean }
 */
export async function searchCommand(query, options = {}) {
  const spinner = ora(chalk.cyan(`Searching for: "${query}"...`)).start();

  const users = await searchUsers(query);
  
  if (users.length === 0) {
    spinner.warn(chalk.yellow('No users found matching that query.'));
    return;
  }
  spinner.succeed(chalk.cyan(`Found ${users.length} user(s).`));

  if (users.length === 1) {
    // Single result: show detail and optionally edit
    await showUserDetail(users[0].userPrincipalName || users[0].id, options.edit);
    return;
  }

  // Multiple results: show list and let user pick
  console.log(chalk.cyan(`Found ${users.length} user(s):\n`));

  const choices = users.map((u) => ({
    name: `${u.displayName} <${u.userPrincipalName || u.mail || u.id}>${u.jobTitle ? ` — ${u.jobTitle}` : ''}${u.department ? ` (${u.department})` : ''}${u.accountEnabled === false ? chalk.red(' [DISABLED]') : ''}`,
    value: u.userPrincipalName || u.id,
  }));

  choices.push({ name: chalk.gray('(cancel)'), value: null });

  const { selected } = await inquirer.prompt([
    {
      type: 'select',
      name: 'selected',
      message: 'Select a user:',
      choices,
      pageSize: 15,
    },
  ]);

  if (!selected) return;

  await showUserDetail(selected, options.edit);
}

/**
 * Display full user detail and optionally launch the editor.
 */
async function showUserDetail(identifier, autoEdit = false) {
  const user = await getUser(identifier);
  if (!user) {
    console.log(chalk.red(`User not found: ${identifier}`));
    return;
  }

  const manager = await getManager(identifier).catch(() => null);

  printUserCard(user, manager);

  if (autoEdit) {
    await editUser(identifier);
    return;
  }

  const { action } = await inquirer.prompt([
    {
      type: 'select',
      name: 'action',
      message: 'What would you like to do?',
      choices: [
        { name: 'Edit this user', value: 'edit' },
        { name: 'Nothing (exit)', value: 'exit' },
      ],
    },
  ]);

  if (action === 'edit') {
    await editUser(identifier);
  }
}

/**
 * Pretty-print a user card to the console.
 */
export function printUserCard(user, manager) {
  const line = chalk.cyan('─'.repeat(52));
  console.log(`\n${line}`);
  console.log(chalk.bold(`  ${user.displayName}`));
  if (user.jobTitle) console.log(`  ${chalk.gray('Title:')}     ${user.jobTitle}`);
  if (user.department) console.log(`  ${chalk.gray('Dept:')}      ${user.department}`);
  console.log(`  ${chalk.gray('UPN:')}       ${user.userPrincipalName}`);
  if (user.mail && user.mail !== user.userPrincipalName) {
    console.log(`  ${chalk.gray('Mail:')}      ${user.mail}`);
  }
  if (user.mobilePhone) console.log(`  ${chalk.gray('Mobile:')}    ${user.mobilePhone}`);
  if (user.businessPhones && user.businessPhones.length > 0) {
    console.log(`  ${chalk.gray('Phone:')}     ${user.businessPhones.join(', ')}`);
  }
  if (user.officeLocation) console.log(`  ${chalk.gray('Office:')}    ${user.officeLocation}`);
  if (user.city || user.country) {
    console.log(`  ${chalk.gray('Location:')}  ${[user.city, user.state, user.country].filter(Boolean).join(', ')}`);
  }
  if (user.usageLocation) console.log(`  ${chalk.gray('Usage loc:')} ${user.usageLocation}`);
  console.log(`  ${chalk.gray('Enabled:')}   ${user.accountEnabled ? chalk.green('Yes') : chalk.red('No')}`);
  if (manager) {
    console.log(`  ${chalk.gray('Manager:')}   ${manager.displayName} <${manager.userPrincipalName}>`);
  }
  console.log(`  ${chalk.gray('ID:')}        ${user.id}`);
  console.log(line + '\n');
}
