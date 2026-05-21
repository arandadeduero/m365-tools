import chalk from 'chalk';
import { parseCSV, rowToUserPayload, stripPasswordFromPayload } from '../utils/csv.js';
import { createUser, updateUser, findUserByUpn, setManager } from '../graph.js';

/**
 * Import users from a CSV file. Creates or updates users (upsert).
 * After creating/updating, assigns the manager if provided.
 *
 * @param {string} filePath - Path to the CSV file
 * @param {object} options - Command options
 */
export async function importUsers(filePath, options = {}) {
  console.log(chalk.cyan(`\nParsing CSV: ${filePath}\n`));

  const { rows, errors } = parseCSV(filePath);

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

  for (const row of rows) {
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
        await updateUser(upn, updatePayload);
        userId = existing.id;
        results.updated.push(upn);
        console.log(chalk.blue(`  [UPDATE] ${upn} — ${userPayload.displayName}`));
      } else {
        // Create new user
        const created = await createUser(userPayload);
        userId = created.id;
        results.created.push(upn);
        console.log(chalk.green(`  [CREATE] ${upn} — ${userPayload.displayName}`));
      }

      // Assign manager if provided
      if (manager) {
        try {
          await setManager(userId, manager);
          console.log(chalk.gray(`           Manager set: ${manager}`));
        } catch (managerErr) {
          results.managerErrors.push({ upn, manager, error: managerErr.message });
          console.log(chalk.yellow(`           Warning: Could not set manager ${manager}: ${managerErr.message}`));
        }
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
}
