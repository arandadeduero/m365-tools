#!/usr/bin/env node

import { program } from 'commander';
import { readFile, access } from 'node:fs/promises';
import { resolve } from 'node:path';
import chalk from 'chalk';
import { initAuth, login, logout, getDomain } from './auth.js';
import { normalizeUpn } from './utils/upn.js';
import { importUsers } from './commands/import.js';
import { searchCommand } from './commands/search.js';
import { editUser } from './commands/edit.js';
import { listCommand } from './commands/list.js';
import { assignManagerByJobTitle } from './commands/assign-manager.js';
import { fixJobTitles, fixOrgChart } from './commands/fix.js';
import { validateCommand } from './commands/validate.js';
import { syncCheckCommand } from './commands/sync-check.js';
import { validateUsersCommand } from './commands/validate-users.js';
import { syncEmployeeIdsCommand } from './commands/sync-employee-ids.js';
import { resetPasswordCommand } from './commands/reset-password.js';
import { REQUIRED_DOMAIN } from './constants.js';
import pkg from '../package.json' with { type: 'json' };

// ---------------------------------------------------------------------------
// Load config
// ---------------------------------------------------------------------------

async function loadConfig(configPath) {
  const fullPath = resolve(configPath);
  try {
    await access(fullPath);
  } catch {
    console.error(chalk.red(`Config file not found: ${fullPath}`));
    console.error(chalk.gray('  Create a config.json with tenantId and clientId.'));
    console.error(chalk.gray('  See config.example.json in the project root for a template.'));
    process.exit(1);
  }
  try {
    const config = JSON.parse(await readFile(fullPath, 'utf8'));
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

program
  .name('m365-users')
  .description('Microsoft 365 user management CLI via Graph API')
  .version(pkg.version)
  .option('-c, --config <path>', 'Path to config.json', resolve(import.meta.dirname, '../config.json'));

// ---------------------------------------------------------------------------
// login command
// ---------------------------------------------------------------------------

program
  .command('login')
  .description('Authenticate against Microsoft 365 (Device Code Flow). Only needed once.')
  .action(async () => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
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
  .command('import <excelFile>')
  .description(
    'Batch import users from an Excel file. Creates new users and updates existing ones.\n' +
    '  Required columns: userPrincipalName, displayName, mailNickname\n' +
    '  Optional columns: givenName, surname, jobTitle, department, manager, ...'
  )
  .option('--limit <number>', 'Limit the number of users to process', parseInt)
  .action(async (excelFile, options) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await importUsers(resolve(excelFile), options);
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
    const config = await loadConfig(opts.config);
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
   .option('--never-signed-in', 'Only show users who have never signed in')
   .option('--for-mailing', 'List all users with actual license, joined by ";"')
   .option('-e, --edit', 'Select a user from the list to edit interactively')
   .action(async (options) => {
     const opts = program.opts();
     const config = await loadConfig(opts.config);
     initAuth(config);
     try {
       await listCommand({
         disabled:      !!options.disabled,
         noManager:     !!options.missingManager,
         noDepartment:  !!options.missingDepartment,
         noJobTitle:    !!options.missingJobTitle,
         neverSignedIn: !!options.neverSignedIn,
         forMailing:    !!options.forMailing,
         edit:          !!options.edit,
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
    const config = await loadConfig(opts.config);
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
    const config = await loadConfig(opts.config);
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
    const config = await loadConfig(opts.config);
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
    const config = await loadConfig(opts.config);
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
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await fixOrgChart();
    } catch (err) {
      handleError(err);
    }
  });

fixCmd
  .command('timezone')
  .description(
    'Set mailbox timezone to Romance Standard Time (Europe/Madrid) for all active users.\n' +
    '  Requires MailboxSettings.ReadWrite permission on the app registration.'
  )
  .option('--dry-run', 'Preview changes without applying them')
  .option('--debug',   'Print raw mailboxSettings API response for each user')
  .action(async (options) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    const { fixTimezone } = await import('./commands/fix-timezone.js');
    try {
      await fixTimezone({ dryRun: !!options.dryRun, debug: !!options.debug });
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// validate-users command
// ---------------------------------------------------------------------------

program
  .command('validate-users')
  .description(
    'Lista todos los usuarios de Microsoft 365 con sus grupos asignados.\n' +
    '  Muestra nombre, UPN, estado (activo/desactivado) y grupos de cada usuario.'
  )
  .option('--empty-groups', 'Mostrar solo los usuarios que no pertenecen a ningún grupo')
  .option('--fix-groups-by-manager', 'Interactivo: elige un manager, luego un grupo, y añade todos sus subordinados a ese grupo')
  .action(async (options) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await validateUsersCommand({
        empty: !!options.emptyGroups,
        fixGroupsByManager: !!options.fixGroupsByManager,
      });
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// sync-check command
// ---------------------------------------------------------------------------

program
  .command('sync-check <excelFile>')
  .description(
    'Cross-reference the employee Excel file with Microsoft 365 and report sync issues.\n' +
    `  Check 1: Active employees (${REQUIRED_DOMAIN}, no Fecha de Baja) must have an M365 account.\n` +
    '  Check 2: Employees with a Fecha de Baja must have their cloud account disabled.\n' +
    '  Use --disable-left-workers to disable and unlicense Check 2 accounts (with double confirmation).'
  )
  .option('--disable-left-workers', 'Disable and remove licences from accounts that should be disabled (Check 2)')
  .action(async (excelFile, options) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await syncCheckCommand(resolve(excelFile), { fix: !!options.disableLeftWorkers });
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// sync-employee-ids command
// ---------------------------------------------------------------------------

program
  .command('sync-employee-ids <excelFile>')
  .description(
    'Lee id_empleado del Excel y lo escribe en el campo employeeId de cada usuario en M365.\n' +
    '  El Excel es la fuente de verdad: siempre sobreescribe el valor en la nube.\n' +
    '  La coincidencia se hace por email (e_mail → userPrincipalName).'
  )
  .action(async (excelFile) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await syncEmployeeIdsCommand(resolve(excelFile));
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// validate command
// ---------------------------------------------------------------------------

program
  .command('validate <excelFile>')
  .description(
    'Validate the employee Excel file against a set of data quality rules.\n' +
    '  Checks: id_empleado (required, unique), e_mail (required, correct domain),\n' +
    '          id_responsable (must reference a known id_empleado).'
  )
  .action(async (excelFile) => {
    try {
      await validateCommand(resolve(excelFile));
    } catch (err) {
      handleError(err);
    }
  });

// ---------------------------------------------------------------------------
// reset-password command
// ---------------------------------------------------------------------------

program
  .command('reset-password <username>')
  .description('Reset a user\'s password and generate a mailto: link with the new credentials.')
  .action(async (username) => {
    const opts = program.opts();
    const config = await loadConfig(opts.config);
    initAuth(config);
    try {
      await resetPasswordCommand(normalizeUpn(username, getDomain()));
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
