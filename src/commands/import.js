import chalk from 'chalk';
import ora from 'ora';
import { parseExcel, rowToUserPayload, stripPasswordFromPayload } from '../utils/excel-importer.js';
import { readExcel } from '../utils/excel.js';
import { createUser, updateUser, findUserByUpn, setManager, listAllUsers, getUserLicenses, getManager } from '../graph.js';
import { getDomain } from '../auth.js';
import { generatePassword } from './reset-password.js';
import { syncEmployeeIds } from './sync-employee-ids.js';

/**
 * Helper to identify changes between payload and existing user.
 */
function getChanges(payload, existing) {
  const changes = [];
  const fieldsToCheck = [
    'displayName', 'givenName', 'surname', 'jobTitle', 'department',
    'mobilePhone', 'officeLocation', 'city', 'state', 'postalCode',
    'country', 'usageLocation', 'preferredLanguage'
  ];

  for (const field of fieldsToCheck) {
    const pValue = payload[field];
    const eValue = existing[field];
    if (pValue !== undefined && pValue !== eValue) {
      changes.push({
        field,
        old: eValue ?? '(null)',
        new: pValue
      });
    }
  }
  return changes;
}

/**
 * Import users from a CSV file. Creates or updates users (upsert).
 * After creating/updating, assigns the manager if provided.
 *
 * @param {string} filePath - Absolute path to the Excel file
 * @param {object} options - Command options
 */
export async function importUsers(filePath, options = {}) {
  const domain = getDomain();
  console.log(chalk.cyan(`\nParsing Excel: ${filePath}\n`));

  const { rows, errors } = await parseExcel(filePath, domain);

  const userSpinner = ora(chalk.cyan('Loading M365 user list to resolve managers...')).start();
  const allUsers = await listAllUsers();
  userSpinner.succeed(chalk.cyan(`Loaded ${allUsers.length} users.`));

  // Create a map of EmployeeID -> M365 User Object for fast lookup
  const employeeIdToUserMap = new Map();
  for (const user of allUsers) {
    if (user.employeeId) {
      employeeIdToUserMap.set(user.employeeId, user);
    }
  }

  // Report parse/validation errors
  if (errors.length > 0) {
    console.log(chalk.yellow(`Found ${errors.length} row(s) with validation errors:\n`));
    for (const e of errors) {
      console.log(chalk.red(`  Row ${e.row} (${e.upn}): ${e.errors.join(', ')}`));
    }
    console.log();
  }

  if (rows.length === 0) {
    console.log(chalk.yellow('No valid rows to process.'));
    return;
  }

  console.log(chalk.cyan(`Processing ${rows.length} user(s)...\n`));

  const results = {
    created: [],
    updated: [],
    failed: [],
    managerErrors: [],
  };

  let processedCount = 0;
  for (const row of rows) {
    if (options.limit && processedCount >= options.limit) {
      console.log(chalk.yellow(`\nReached limit of ${options.limit} users. Stopping.`));
      break;
    }
    processedCount++;
    const { userPayload, manager } = rowToUserPayload(row);
    const upn = userPayload.userPrincipalName;

    try {
      // Check if user already exists
      const existing = await findUserByUpn(upn);

      let userId;

      if (existing) {
        // Update existing user
        const updatePayload = stripPasswordFromPayload(userPayload);
        // Remove fields that cannot be updated via PATCH in the same call
        delete updatePayload.mailNickname; // can cause issues if not changed

        const changes = getChanges(updatePayload, existing);
        const isReady = changes.length === 0;

        if (isReady) {
          console.log(chalk.green(`  [IS-READY] ${upn} — ${userPayload.displayName}`));
          userId = existing.id;
        } else {
          await updateUser(upn, updatePayload);
          userId = existing.id;
          results.updated.push(upn);
          console.log(chalk.blue(`  [UPDATE] ${upn} — ${userPayload.displayName}`));
          for (const change of changes) {
            console.log(chalk.gray(`           Change ${change.field}: ${change.old} -> ${change.new}`));
          }
        }
      } else {
        // Create new user
        if (!userPayload.passwordProfile) {
          userPayload.passwordProfile = {
            password: generatePassword(),
            forceChangePasswordNextSignIn: true,
          };
        }
        const created = await createUser(userPayload);
        userId = created.id;
        results.created.push(upn);
        console.log(chalk.green(`  [CREATE] ${upn} — ${userPayload.displayName}`));
      }

      // Assign manager if provided
      if (manager) {
        let managerUpn = String(manager);
        let managerM365Id = null;
        const isEmployeeId = /^\d+$/.test(managerUpn);

        if (isEmployeeId) {
          const managerUser = employeeIdToUserMap.get(managerUpn);
          managerM365Id = managerUser?.id;
          if (!managerM365Id) {
            results.managerErrors.push({ upn, manager: managerUpn, error: 'Manager EmployeeID not found in M365' });
            console.log(chalk.yellow(`           Warning: Could not find manager with EmployeeID ${managerUpn}`));
            continue;
          }
        } else {
          // Construct UPN
          managerUpn = `${managerUpn.toLowerCase()}@${domain}`;
        }

        // Fetch current manager to compare
        let currentManager = null;
        try {
          currentManager = await getManager(userId);
        } catch (e) {
          // ignore
        }

        // Check if manager is already set correctly
        const isManagerAlreadySet = isEmployeeId
          ? currentManager?.id === managerM365Id
          : currentManager?.userPrincipalName?.toLowerCase() === managerUpn.toLowerCase();

        if (!isManagerAlreadySet) {
          console.log(chalk.gray(`           Manager: Cloud=${currentManager?.userPrincipalName || '(none)'} | Excel=${managerUpn}`));

          try {
            await setManager(userId, managerUpn, managerM365Id);
            console.log(chalk.gray(`           Manager set: ${managerUpn} (resolved to ${managerM365Id || 'UPN'})`));
          } catch (managerErr) {
            results.managerErrors.push({ upn, manager: managerUpn, error: managerErr.message });
            console.log(chalk.yellow(`           Warning: Could not set manager ${managerUpn}: ${managerErr.message}`));
          }
        }
      }


      // Verify license
      const licenses = await getUserLicenses(userId);
      if (licenses.length === 0) {
        console.log(chalk.yellow(`           Warning: User ${upn} has no active licenses.`));
      }
    } catch (err) {
      results.failed.push({ upn, error: err.message || String(err) });
      console.log(chalk.red(`  [FAILED] ${upn}: ${err.message || err}`));
    }
  }

  // Summary
  console.log(chalk.cyan('\n----------------------------------------'));
  console.log(chalk.bold('Import Summary:'));
  console.log(`  ${chalk.green(`Created: ${results.created.length}`)}`);
  console.log(`  ${chalk.blue(`Updated: ${results.updated.length}`)}`);
  console.log(`  ${chalk.red(`Failed:  ${results.failed.length}`)}`);
  if (results.managerErrors.length > 0) {
    console.log(`  ${chalk.yellow(`Manager errors: ${results.managerErrors.length}`)}`);
  }
  console.log(chalk.cyan('----------------------------------------\n'));

  if (results.failed.length > 0) {
    console.log(chalk.red('Failed users:'));
    for (const f of results.failed) {
      console.log(chalk.red(`  - ${f.upn}: ${f.error}`));
    }
    console.log();
  }

  if (results.managerErrors.length > 0) {
    console.log(chalk.yellow('Manager assignment errors:'));
    for (const m of results.managerErrors) {
      console.log(chalk.yellow(`  - ${m.upn} -> ${m.manager}: ${m.error}`));
    }
    console.log();
  }

  // ── Sync employee IDs & types ─────────────────────────────────────────────
  // Re-read raw Excel rows (the import normalizer transforms them; sync needs originals)
  // and re-fetch cloud users so we compare against the freshly-imported state.
  console.log(chalk.cyan('----------------------------------------'));
  console.log(chalk.bold('Syncing employee IDs & types from Excel…\n'));

  let rawRows;
  try {
    ({ rows: rawRows } = await readExcel(filePath));
  } catch (err) {
    console.log(chalk.yellow(`  Warning: Could not re-read Excel for employee ID sync: ${err.message}`));
    return;
  }

  // Refresh cloud user list to include any users just created/updated
  let freshCloudUsers;
  try {
    freshCloudUsers = await listAllUsers({ onProgress: () => {} });
  } catch (err) {
    console.log(chalk.yellow(`  Warning: Could not fetch cloud users for employee ID sync: ${err.message}`));
    return;
  }

  const cloudMap = new Map(
    freshCloudUsers.map((u) => [u.userPrincipalName.toLowerCase(), u]),
  );

  await syncEmployeeIds(rawRows, cloudMap);
  console.log(chalk.cyan('----------------------------------------\n'));
}
