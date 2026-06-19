/**
 * validate-users command
 *
 * Fetches all users from Microsoft 365 and, for each one, lists the groups
 * they belong to. Uses $batch requests (20 per batch) to minimise API round
 * trips when the tenant has many users.
 *
 * Output is a sorted table:
 *   Nombre  |  UPN  |  Estado  |  Manager  |  Grupos
 *
 * Options:
 *   --empty-groups           Show only users with no groups
 *   --fix-groups-by-manager  Interactive: pick a manager → find all their
 *                            reports → pick a group → add them all
 *
 * Usage:
 *   m365-users validate-users
 *   m365-users validate-users --empty-groups
 *   m365-users validate-users --fix-groups-by-manager
 */

import chalk from 'chalk';
import inquirer from 'inquirer';
import ora from 'ora';
import { listAllUsers, fetchGroupsMap, searchUsers, searchGroups, addMemberToGroup } from '../graph.js';
import { filterAutoGroups } from '../utils/cache.js';

const filterGroups = filterAutoGroups;

// ---------------------------------------------------------------------------
// Column widths
// ---------------------------------------------------------------------------
const NAME_W    = 40;
const UPN_W     = 46;
const STATUS_W  = 12;
const MANAGER_W = 35;

// ---------------------------------------------------------------------------
// Print table helpers
// ---------------------------------------------------------------------------

function printHeader() {
  console.log(
    chalk.bold(
      'Nombre'.padEnd(NAME_W) + '  ' +
      'UPN'.padEnd(UPN_W) + '  ' +
      'Estado'.padEnd(STATUS_W) + '  ' +
      'Manager'.padEnd(MANAGER_W) + '  ' +
      'Grupos',
    ),
  );
  console.log(chalk.gray('─'.repeat(NAME_W + UPN_W + STATUS_W + MANAGER_W + 80)));
}

function printUserRow(user, groups) {
  const name    = (user.displayName ?? '').slice(0, NAME_W - 1).padEnd(NAME_W);
  const upn     = (user.userPrincipalName ?? '').slice(0, UPN_W - 1).padEnd(UPN_W);
  const status  = user.accountEnabled
    ? chalk.green('activo'.padEnd(STATUS_W))
    : chalk.red('desactivado'.padEnd(STATUS_W));
  const manager = (user.manager?.displayName ?? user.manager?.upn ?? '')
    .slice(0, MANAGER_W - 1).padEnd(MANAGER_W);
  const groupLabel = groups.length === 0
    ? chalk.gray('(sin grupos)')
    : groups
        .slice()
        .sort((a, b) => a.displayName.localeCompare(b.displayName, 'es'))
        .map((g) => chalk.cyan(g.displayName))
        .join(', ');

  console.log(`${chalk.white(name)}  ${chalk.gray(upn)}  ${status}  ${chalk.yellow(manager)}  ${groupLabel}`);
}

// ---------------------------------------------------------------------------
// --fix-groups-by-manager flow
// ---------------------------------------------------------------------------

async function runFixGroupsByManager(users, groupsMap) {
  // ── Step 1: pick a manager ────────────────────────────────────────────────
  const { managerQuery } = await inquirer.prompt([{
    type: 'input',
    name: 'managerQuery',
    message: 'Buscar manager (nombre o email):',
    validate: (v) => v.trim().length >= 2 || 'Introduce al menos 2 caracteres',
  }]);

  const managerResults = await searchUsers(managerQuery.trim());
  if (managerResults.length === 0) {
    console.log(chalk.red('No se encontró ningún usuario con esa búsqueda.'));
    return;
  }

  const { managerId } = await inquirer.prompt([{
    type: 'select',
    name: 'managerId',
    message: 'Selecciona el manager:',
    choices: managerResults.map((u) => ({
      name: `${u.displayName}  ${chalk.gray(u.userPrincipalName)}`,
      value: u.id,
    })),
  }]);

  const selectedManager = managerResults.find((u) => u.id === managerId);

  // ── Step 2: find all users whose manager is the selected one ──────────────
  const reports = users.filter((u) => u.manager && u.id === managerId
    ? false  // skip the manager themselves
    : u.manager?.upn?.toLowerCase() === selectedManager.userPrincipalName.toLowerCase()
      || u.manager?.id === managerId,
  );

  if (reports.length === 0) {
    console.log(chalk.yellow(`\nNo se encontraron usuarios con manager "${selectedManager.displayName}".`));
    return;
  }

  console.log(chalk.gray(`\nUsuarios a cargo de ${chalk.bold(selectedManager.displayName)} (${reports.length}):\n`));
  printHeader();
  for (const u of reports) {
    printUserRow(u, filterGroups(groupsMap.get(u.id) ?? []));
  }
  console.log('');

  // ── Step 3: pick a group ──────────────────────────────────────────────────
  const { groupQuery } = await inquirer.prompt([{
    type: 'input',
    name: 'groupQuery',
    message: 'Buscar grupo (nombre parcial):',
    validate: (v) => v.trim().length >= 2 || 'Introduce al menos 2 caracteres',
  }]);

  let groupResults;
  try {
    groupResults = await searchGroups(groupQuery.trim());
  } catch (err) {
    console.error(chalk.red(`Error buscando grupos: ${err.message}`));
    return;
  }

  if (groupResults.length === 0) {
    console.log(chalk.red('No se encontró ningún grupo con ese nombre.'));
    return;
  }

  const { groupId } = await inquirer.prompt([{
    type: 'select',
    name: 'groupId',
    message: 'Selecciona el grupo:',
    choices: groupResults.map((g) => ({
      name: g.displayName,
      value: g.id,
    })),
  }]);

  const selectedGroup = groupResults.find((g) => g.id === groupId);

  // ── Step 4: filter out users already in the group ─────────────────────────
  const toAdd = reports.filter((u) => {
    const userGroups = groupsMap.get(u.id) ?? [];
    return !userGroups.some((g) => g.id === groupId);
  });

  const alreadyIn = reports.length - toAdd.length;

  if (toAdd.length === 0) {
    console.log(chalk.green(`\n✔  Todos los usuarios ya pertenecen al grupo "${selectedGroup.displayName}".`));
    return;
  }

  // ── Step 5: preview ───────────────────────────────────────────────────────
  console.log('');
  console.log(
    chalk.bold(`Se añadirán ${toAdd.length} usuario(s) al grupo `) +
    chalk.cyan.bold(selectedGroup.displayName) +
    (alreadyIn > 0 ? chalk.gray(` (${alreadyIn} ya son miembros, se omiten)`) : '') +
    ':',
  );
  console.log('');

  for (const u of toAdd) {
    console.log(
      `  ${chalk.red('►')} ${chalk.white((u.displayName ?? '').padEnd(NAME_W))}  ${chalk.gray(u.userPrincipalName)}`,
    );
  }
  console.log('');

  // ── Step 6: double confirmation ───────────────────────────────────────────
  const { confirm1 } = await inquirer.prompt([{
    type: 'confirm',
    name: 'confirm1',
    message: chalk.yellow(`¿Añadir ${toAdd.length} usuario(s) al grupo "${selectedGroup.displayName}"?`),
    default: false,
  }]);
  if (!confirm1) { console.log(chalk.gray('\nCancelado — sin cambios.\n')); return; }

  const { confirm2 } = await inquirer.prompt([{
    type: 'confirm',
    name: 'confirm2',
    message: chalk.red(`CONFIRMACIÓN FINAL: añadir ${toAdd.length} usuario(s) a "${selectedGroup.displayName}"?`),
    default: false,
  }]);
  if (!confirm2) { console.log(chalk.gray('\nCancelado — sin cambios.\n')); return; }

  // ── Step 7: apply ─────────────────────────────────────────────────────────
  console.log('');
  let ok = 0;
  let failed = 0;

  for (const u of toAdd) {
    process.stdout.write(`  ${chalk.cyan((u.displayName ?? '').padEnd(NAME_W))}  `);
    try {
      await addMemberToGroup(groupId, u.id);
      console.log(chalk.green('✔  añadido'));
      ok++;
    } catch (err) {
      console.log(chalk.red(`✖  ${err.message}`));
      failed++;
    }
  }

  console.log('');
  if (failed === 0) {
    console.log(chalk.green.bold(`✔  ${ok} usuario(s) añadidos al grupo "${selectedGroup.displayName}".`));
  } else {
    console.log(chalk.yellow(`  ${ok} correctos`) + '  ' + chalk.red(`${failed} fallidos`));
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

export async function validateUsersCommand({ empty = false, fixGroupsByManager = false } = {}) {
  // ── 1. Fetch all users ────────────────────────────────────────────────────
  const userSpinner = ora(chalk.cyan('Fetching users from Microsoft 365…')).start();

  let users;
  try {
    users = await listAllUsers({ checkManager: true, onProgress: () => {} });
  } catch (err) {
    userSpinner.fail(chalk.red(`\nFailed to fetch users: ${err.message}`));
    process.exit(1);
  }

  users.sort((a, b) => (a.displayName ?? '').localeCompare(b.displayName ?? '', 'es'));
  userSpinner.succeed(chalk.cyan(`Found ${users.length} user(s).`));

  // ── 2. Fetch group memberships in batch ───────────────────────────────────
  const groupSpinner = ora(chalk.cyan('Fetching group memberships…')).start();
  let groupsMap;
  try {
    groupsMap = await fetchGroupsMap(users);
    groupSpinner.succeed(chalk.cyan('Fetched group memberships.'));
  } catch (err) {
    groupSpinner.fail(chalk.red(`\nFailed to fetch group memberships: ${err.message}`));
    process.exit(1);
  }

  // ── 3. Fix mode ───────────────────────────────────────────────────────────
  if (fixGroupsByManager) {
    await runFixGroupsByManager(users, groupsMap);
    return;
  }

  // ── 4. Print report ────────────────────────────────────────────────────────
  const usersToShow = empty
    ? users.filter((u) => filterGroups(groupsMap.get(u.id) ?? []).length === 0)
    : users;

  if (empty && usersToShow.length === 0) {
    console.log(chalk.green('✔  Todos los usuarios pertenecen a al menos un grupo.'));
    process.exit(0);
  }

  printHeader();
  for (const user of usersToShow) {
    printUserRow(user, filterGroups(groupsMap.get(user.id) ?? []));
  }

  // ── 5. Summary ────────────────────────────────────────────────────────────
  const activeCount   = users.filter((u) => u.accountEnabled).length;
  const disabledCount = users.length - activeCount;
  const noGroupCount  = users.filter((u) => filterGroups(groupsMap.get(u.id) ?? []).length === 0).length;

  console.log(chalk.gray('\n' + '─'.repeat(NAME_W + UPN_W + STATUS_W + MANAGER_W + 80)));
  console.log(
    chalk.bold(`Total: ${users.length} usuario(s)`) +
    chalk.gray('  |  ') +
    chalk.green(`${activeCount} activo(s)`) +
    chalk.gray('  |  ') +
    chalk.red(`${disabledCount} desactivado(s)`) +
    chalk.gray('  |  ') +
    chalk.yellow(`${noGroupCount} sin grupos`),
  );
}
