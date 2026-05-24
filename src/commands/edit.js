import chalk from 'chalk';
import inquirer from 'inquirer';
import { getUser, updateUser, setManager, getManager, findUserByUpn, getUserGroups, listAllGroups, addMemberToGroup, removeMemberFromGroup } from '../graph.js';
import { printUserCard } from './search.js';
import { getDomain } from '../auth.js';
import { normalizeUpn } from '../utils/upn.js';

// All editable fields grouped by category
export const FIELD_GROUPS = [
  {
    name: 'Identity',
    fields: [
      { name: 'displayName', message: 'Display Name', type: 'input' },
      { name: 'givenName', message: 'First Name', type: 'input' },
      { name: 'surname', message: 'Last Name', type: 'input' },
      { name: 'mailNickname', message: 'Mail Nickname (alias)', type: 'input' },
    ],
  },
  {
    name: 'Job Information',
    fields: [
      { name: 'jobTitle', message: 'Job Title', type: 'input' },
      { name: 'department', message: 'Department', type: 'input' },
      { name: 'companyName', message: 'Company Name', type: 'input' },
      { name: 'employeeId', message: 'Employee ID', type: 'input' },
      { name: 'employeeType', message: 'Employee Type (Employee/Contractor/Consultant/Vendor)', type: 'input' },
      { name: 'employeeHireDate', message: 'Hire Date (YYYY-MM-DD)', type: 'input' },
    ],
  },
  {
    name: 'Contact',
    fields: [
      { name: 'mobilePhone', message: 'Mobile Phone', type: 'input' },
      { name: 'businessPhone', message: 'Business Phone', type: 'input', graphField: 'businessPhones' },
      { name: 'officeLocation', message: 'Office Location', type: 'input' },
    ],
  },
  {
    name: 'Address',
    fields: [
      { name: 'streetAddress', message: 'Street Address', type: 'input' },
      { name: 'city', message: 'City', type: 'input' },
      { name: 'state', message: 'State / Province', type: 'input' },
      { name: 'postalCode', message: 'Postal Code', type: 'input' },
      { name: 'country', message: 'Country (e.g. ES, US)', type: 'input' },
    ],
  },
  {
    name: 'Settings',
    fields: [
      { name: 'usageLocation', message: 'Usage Location (ISO 3166 2-letter, e.g. ES)', type: 'input' },
      { name: 'preferredLanguage', message: 'Preferred Language (e.g. en-US, es-ES)', type: 'input' },
      {
        name: 'accountEnabled',
        message: 'Account Enabled',
        type: 'confirm',
        isBoolean: true,
      },
    ],
  },
  {
    name: 'Manager (Org Chart)',
    fields: [
      { name: 'manager', message: 'Manager UPN (email)', type: 'input', isManager: true },
    ],
  },
  {
    name: 'Password',
    fields: [
      { name: 'newPassword', message: 'New Password (leave blank to skip)', type: 'password', isPassword: true },
    ],
  },
  {
    name: 'Groups',
    fields: [
      { name: 'groups', message: 'Group memberships', type: 'checkbox', isGroups: true },
    ],
  },
];

/**
 * Interactively edit a user by UPN or object ID.
 * @param {string} identifier - UPN or object ID
 */
export async function editUser(identifier) {
  const user = await getUser(identifier);
  if (!user) {
    console.log(chalk.red(`User not found: ${identifier}`));
    return;
  }

  const currentManager = await getManager(identifier).catch(() => null);

  printUserCard(user, currentManager);
  console.log(chalk.cyan('Editing user. Leave fields blank to keep current values.\n'));

  // Ask which section to edit
  const groupChoices = FIELD_GROUPS.map((g) => ({
    name: g.name,
    value: g.name,
  }));
  groupChoices.push({ name: chalk.bold('Save all changes'), value: '__SAVE__' });
  groupChoices.push({ name: chalk.gray('Cancel (discard)'), value: '__CANCEL__' });

  const pendingChanges = {};
  let pendingManager = null;
  let pendingPassword = null;
  let pendingGroups = null; // { toAdd: [{id,displayName}], toRemove: [{id,displayName}] }

  // Interactive edit loop
  while (true) {
    showPendingChanges(pendingChanges, pendingManager, pendingPassword, pendingGroups);

    const { section } = await inquirer.prompt([
      {
        type: 'list',
        name: 'section',
        message: 'Select section to edit:',
        choices: groupChoices,
        pageSize: 15,
      },
    ]);

    if (section === '__CANCEL__') {
      console.log(chalk.gray('\nEdit cancelled. No changes made.\n'));
      return;
    }

    if (section === '__SAVE__') {
      break;
    }

    const group = FIELD_GROUPS.find((g) => g.name === section);
    await editSection(group, user, currentManager, pendingChanges, (m) => { pendingManager = m; }, (p) => { pendingPassword = p; }, (g) => { pendingGroups = g; }, identifier);
  }

  // Apply changes
  if (Object.keys(pendingChanges).length === 0 && pendingManager === null && pendingPassword === null && pendingGroups === null) {
    console.log(chalk.yellow('\nNo changes to save.\n'));
    return;
  }

  const { confirm } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirm',
      message: `Apply ${Object.keys(pendingChanges).length + (pendingManager ? 1 : 0) + (pendingPassword ? 1 : 0) + (pendingGroups ? pendingGroups.toAdd.length + pendingGroups.toRemove.length : 0)} change(s) to ${user.displayName}?`,
      default: true,
    },
  ]);

  if (!confirm) {
    console.log(chalk.gray('\nCancelled. No changes saved.\n'));
    return;
  }

  // Build update payload
  const updatePayload = { ...pendingChanges };

  if (pendingPassword) {
    updatePayload.passwordProfile = {
      password: pendingPassword,
      forceChangePasswordNextSignIn: true,
    };
  }

  // Apply user property updates
  if (Object.keys(updatePayload).length > 0) {
    try {
      await updateUser(identifier, updatePayload);
      console.log(chalk.green('\nUser properties updated successfully.'));
    } catch (err) {
      console.log(chalk.red(`\nFailed to update user: ${err.message || err}`));
      return;
    }
  }

  // Apply manager update
  if (pendingManager !== null) {
    if (pendingManager === '') {
      // Clear manager is not directly supported via Graph in most tenants, warn user
      console.log(chalk.yellow('\nNote: Clearing the manager is not supported via this tool. Use the Azure portal.'));
    } else {
      try {
        await setManager(identifier, pendingManager);
        console.log(chalk.green(`Manager set to: ${pendingManager}`));
      } catch (err) {
        console.log(chalk.red(`Failed to set manager: ${err.message || err}`));
      }
    }
  }

  // Apply group changes
  if (pendingGroups !== null && (pendingGroups.toAdd.length > 0 || pendingGroups.toRemove.length > 0)) {
    await applyGroupChanges(user.id, pendingGroups);
  }

  console.log(chalk.cyan('\nDone.\n'));
}

/**
 * Apply pending group changes (add/remove) for a user.
 * @param {string} userId - Azure AD object ID (not UPN)
 */
async function applyGroupChanges(userId, pendingGroups) {
  for (const g of pendingGroups.toAdd) {
    try {
      await addMemberToGroup(g.id, userId);
      console.log(chalk.green(`  Added to group:     ${g.displayName}`));
    } catch (err) {
      console.log(chalk.red(`  Failed to add to ${g.displayName}: ${err.message}`));
    }
  }
  for (const g of pendingGroups.toRemove) {
    try {
      await removeMemberFromGroup(g.id, userId);
      console.log(chalk.green(`  Removed from group: ${g.displayName}`));
    } catch (err) {
      console.log(chalk.red(`  Failed to remove from ${g.displayName}: ${err.message}`));
    }
  }
}

/**
 * Edit a specific section of user properties.
 */
export async function editSection(group, user, currentManager, pendingChanges, setManager_, setPassword, setGroups, identifier) {
  // ── Groups section — special interactive checkbox flow ──────────────────
  if (group.fields.some((f) => f.isGroups)) {
    console.log(chalk.gray('\nFetching all groups and current memberships…'));

    const [allGroups, currentMemberships] = await Promise.all([
      listAllGroups(),
      getUserGroups(user.id),
    ]);

    const currentIds = new Set(currentMemberships.map((g) => g.id));

    const { selected } = await inquirer.prompt([{
      type: 'checkbox',
      name: 'selected',
      message: 'Select groups (space to toggle, enter to confirm):',
      choices: allGroups.map((g) => ({
        name: currentIds.has(g.id)
          ? `${g.displayName}  ${chalk.green('(miembro)')}`
          : g.displayName,
        value: g.id,
        checked: currentIds.has(g.id),
      })),
      pageSize: 20,
    }]);

    const selectedIds = new Set(selected);

    const toAdd    = allGroups.filter((g) => selectedIds.has(g.id) && !currentIds.has(g.id));
    const toRemove = allGroups.filter((g) => !selectedIds.has(g.id) && currentIds.has(g.id));

    if (toAdd.length === 0 && toRemove.length === 0) {
      console.log(chalk.gray('No group changes.'));
      setGroups(null);
    } else {
      setGroups({ toAdd, toRemove });
    }
    return;
  }
  const questions = group.fields.map((field) => {
    const current = getCurrentValue(field, user, currentManager);
    const defaultVal = current !== null && current !== undefined ? String(current) : '';

    if (field.isPassword) {
      return {
        type: 'password',
        name: field.name,
        message: field.message,
        mask: '*',
      };
    }

    if (field.isBoolean) {
      return {
        type: 'confirm',
        name: field.name,
        message: `${field.message} (current: ${current ? 'Yes' : 'No'})`,
        default: current !== false,
      };
    }

    if (field.isManager) {
      const domain = getDomain();
      return {
        type: 'input',
        name: field.name,
        message: `${field.message}${domain ? chalk.gray(` [@${domain} added if missing]`) : ''}${defaultVal ? chalk.gray(` (current: ${defaultVal})`) : ''}:`,
        default: defaultVal,
        filter: (v) => normalizeUpn(v, domain),
        validate: async (v) => {
          const normalized = normalizeUpn(v, domain);
          // Empty = keep current, skip validation
          if (!normalized || normalized.trim() === '') return true;
          // Same as current = no change, skip validation
          if (normalized === defaultVal) return true;
          process.stdout.write(chalk.gray(' Verifying...'));
          const found = await findUserByUpn(normalized).catch(() => null);
          if (!found) {
            return `User not found: ${normalized}`;
          }
          return true;
        },
      };
    }

    return {
      type: 'input',
      name: field.name,
      message: `${field.message}${defaultVal ? chalk.gray(` (current: ${defaultVal})`) : ''}:`,
      default: defaultVal,
    };
  });

  const answers = await inquirer.prompt(questions);

  for (const field of group.fields) {
    const answer = answers[field.name];

    if (field.isPassword) {
      if (answer && answer.trim() !== '') {
        setPassword(answer.trim());
      }
      continue;
    }

    if (field.isManager) {
      const current = currentManager ? currentManager.userPrincipalName : '';
      if (answer !== current) {
        setManager_(answer.trim());
      }
      continue;
    }

    if (field.isBoolean) {
      const current = getCurrentValue(field, user, currentManager);
      if (answer !== current) {
        pendingChanges[field.name] = answer;
      }
      continue;
    }

    if (field.graphField === 'businessPhones') {
      const current = user.businessPhones && user.businessPhones[0] ? user.businessPhones[0] : '';
      if (answer !== current && answer !== '') {
        pendingChanges.businessPhones = [answer];
      }
      continue;
    }

    const current = getCurrentValue(field, user, currentManager);
    const currentStr = current !== null && current !== undefined ? String(current) : '';
    if (answer !== currentStr && answer !== '') {
      pendingChanges[field.name] = answer;
    }
  }
}

function getCurrentValue(field, user, currentManager) {
  if (field.isManager) {
    return currentManager ? currentManager.userPrincipalName : '';
  }
  if (field.graphField === 'businessPhones') {
    return user.businessPhones && user.businessPhones[0] ? user.businessPhones[0] : '';
  }
  return user[field.name] !== undefined ? user[field.name] : '';
}

function showPendingChanges(pendingChanges, pendingManager, pendingPassword, pendingGroups) {
  const groupCount = pendingGroups ? pendingGroups.toAdd.length + pendingGroups.toRemove.length : 0;
  const count = Object.keys(pendingChanges).length + (pendingManager !== null ? 1 : 0) + (pendingPassword ? 1 : 0) + groupCount;
  if (count === 0) return;

  console.log(chalk.yellow(`\n  Pending changes (${count}):`));
  for (const [key, val] of Object.entries(pendingChanges)) {
    if (key === 'accountEnabled') {
      console.log(chalk.yellow(`    ${key}: ${val ? 'Enabled' : 'Disabled'}`));
    } else if (key === 'businessPhones') {
      console.log(chalk.yellow(`    businessPhone: ${val[0]}`));
    } else {
      console.log(chalk.yellow(`    ${key}: ${val}`));
    }
  }
  if (pendingManager !== null) {
    console.log(chalk.yellow(`    manager: ${pendingManager || '(clear)'}`));
  }
  if (pendingPassword) {
    console.log(chalk.yellow(`    password: (new password set)`));
  }
  if (pendingGroups) {
    for (const g of pendingGroups.toAdd) {
      console.log(chalk.green(`    + group: ${g.displayName}`));
    }
    for (const g of pendingGroups.toRemove) {
      console.log(chalk.red(`    - group: ${g.displayName}`));
    }
  }
  console.log();
}
