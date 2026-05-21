#!/usr/bin/env node

import { program } from 'commander';
import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import { initAuth, login, logout, getDomain } from './auth.js';
import { normalizeUpn } from './utils/upn.js';
import { importUsers } from './commands/import.js';
import { searchCommand } from './commands/search.js';
import { editUser } from './commands/edit.js';
import { listCommand } from './commands/list.js';
import { assignManagerByJobTitle } from './commands/assign-manager.js';
import { fixJobTitles, fixOrgChart } from './commands/fix.js';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// Load config
// ---------------------------------------------------------------------------

function loadConfig(configPath) {
  const fullPath = resolve(configPath);
  if (!existsSync(fullPath)) {
    console.error(chalk.red(`Config file not found: ${fullPath}`));
    console.error(chalk.gray('  Create a config.json with tenantId and clientId.'));
    console.error(chalk.gray('  See config.json in the project root for a template.'));
    process.exit(1);
  }
  try {
    const config = JSON.parse(readFileSync(fullPath, 'utf8'));
    if (!config.tenantId || config.tenantId === 'YOUR_TENANT_ID') {
      console.error(chalk.red('config.json: tenantId is not configured.'));
      process.exit(1);
    }
    if (!config.clientId || config.clientId === 'YOUR_CLIENT_ID') {
      console.error(chalk.red('config.json: clientId is not configured.'));
      process.exit(1);
    }
    return config;
  } catch (err) {
    console.error(chalk.red(`Failed to parse config file: ${err.message}`));
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------
// Global error handler
// ---------------------------------------------------------------------------

function handleError(err) {
  if (err.statusCode) {
    console.error(chalk.red(`\nGraph API error ${err.statusCode}: ${err.message || err.code}`));
    if (err.statusCode === 401) {
      console.error(chalk.yellow('  Your session may have expired. Run: m365-users login'));
    } else if (err.statusCode === 403) {
      console.error(chalk.yellow('  Insufficient permissions. Check that your app has User.ReadWrite.All and Directory.ReadWrite.All.'));
    }
  } else {
    console.error(chalk.red(`\nError: ${err.message || err}`));
  }
  process.exit(1);
}

// ---------------------------------------------------------------------------
// CLI Definition
// ---------------------------------------------------------------------------

const pkg = JSON.parse(readFileSync(resolve(__dirname, '../package.json'), 'utf8'));

program
  .name('m365-users')
  .description('Microsoft 365 user management CLI via Graph API')
  .version(pkg.version)
  .option('-c, --config <path>', 'Path to config.json', resolve(__dirname, '../config.json'));

// ---------------------------------------------------------------------------
// login command
// ---------------------------------------------------------------------------

program
  .command('login')
  .description('Authenticate against Microsoft 365 (Device Code Flow). Only needed once.')
  .action(async () => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await login();
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// logout command
// ---------------------------------------------------------------------------

program
  .command('logout')
  .description('Clear the saved authentication session.')
  .action(async () => {
    await logout();
  });

// ---------------------------------------------------------------------------
// import command
// ---------------------------------------------------------------------------

program
  .command('import <csvFile>')
  .description(
    'Batch import users from a CSV file. Creates new users and updates existing ones.\n' +
    '  Required CSV columns: userPrincipalName, displayName, mailNickname, password\n' +
    '  Optional columns: givenName, surname, jobTitle, department, manager, ...'
  )
  .action(async (csvFile, options) => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await importUsers(resolve(csvFile), options);
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// search command
// ---------------------------------------------------------------------------

program
  .command('search <query>')
  .description('Search users by name or email. Use quotes for multi-word names.')
  .option('-e, --edit', 'Open the editor immediately after selecting a user')
  .action(async (query, options) => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await searchCommand(normalizeUpn(query, getDomain()), options);
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// list command
// ---------------------------------------------------------------------------

program
  .command('list')
  .description('List active users in the tenant (default). Use --disabled to list disabled users instead.')
  .option('--disabled', 'List disabled (inactive) users instead of active ones')
  .option('--missing-manager', 'Only show users without a manager assigned')
  .option('--missing-department', 'Only show users without a department')
  .option('--missing-job-title', 'Only show users without a job title')
  .option('-e, --edit', 'Select a user from the list to edit interactively')
  .action(async (options) => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await listCommand({
        disabled:     !!options.disabled,
        noManager:    !!options.missingManager,
        noDepartment: !!options.missingDepartment,
        noJobTitle:   !!options.missingJobTitle,
        edit:         !!options.edit,
      });
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// assign-manager command
// ---------------------------------------------------------------------------

program
  .command('assign-manager')
  .description(
    'Bulk-assign a manager to all active users sharing the same job title.\n' +
    '  Shows a list of all existing job titles to choose from — no manual typing.'
  )
  .action(async () => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await assignManagerByJobTitle();
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// edit command
// ---------------------------------------------------------------------------

program
  .command('edit <upnOrId>')
  .description('Interactively edit a user by their UPN (email) or Azure AD object ID.')
  .action(async (upnOrId) => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await editUser(normalizeUpn(upnOrId, getDomain()));
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// whoami command (convenience)
// ---------------------------------------------------------------------------

program
  .command('whoami')
  .description('Show details for a specific user (by UPN or ID).')
  .argument('<upnOrId>', 'User UPN or object ID')
  .action(async (upnOrId) => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    const { getUser, getManager } = await import('./graph.js');
    const { printUserCard } = await import('./commands/search.js');
    const id = normalizeUpn(upnOrId, getDomain());
    try {
      const user = await getUser(id);
      if (!user) {
        console.log(chalk.red(`User not found: ${id}`));
        process.exit(1);
      }
      const manager = await getManager(id).catch(() => null);
      printUserCard(user, manager);
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// fix command
// ---------------------------------------------------------------------------

const fixCmd = program
  .command('fix')
  .description('Data quality fixes.');

fixCmd
  .command('job-titles')
  .description('Normalise all job titles to Sentence case (first letter uppercase, rest lowercase).')
  .action(async () => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await fixJobTitles();
    } catch (err) {
      handleError(err);
    }
  });

fixCmd
  .command('org-chart')
  .description(
    'Verify the org chart integrity: every user must have a manager, ' +
    'all chains must reach a single root, and there must be no cycles.'
  )
  .action(async () => {
    const opts = program.opts();
    const config = loadConfig(opts.config);
    initAuth(config);
    try {
      await fixOrgChart();
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------

program.parse(process.argv);

if (process.argv.length < 3) {
  program.help();
}
