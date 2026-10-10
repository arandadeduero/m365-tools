/**
 * Validation rules for the employee Excel file.
 *
 * Each rule is a plain object:
 * {
 *   id:          string          — unique rule identifier
 *   description: string          — human-readable rule summary
 *   validate:    (rows) => Issue[]
 * }
 *
 * Issue shape:
 * {
 *   row:     number   — 1-based data row number (not counting the header)
 *   column:  string   — column name
 *   value:   string   — the offending value
 *   message: string   — description of the problem
 * }
 *
 * To add a new rule, push a new object into the RULES array below.
 */

import { REQUIRED_DOMAIN } from '../constants.js';

// ---------------------------------------------------------------------------
// Rule: id_empleado — required and unique
// ---------------------------------------------------------------------------

const ruleIdEmpleado = {
  id: 'id_empleado-required-unique',
  description: 'id_empleado must be filled in every row and must be unique',
  validate(rows) {
    const issues = [];
    const seen = new Map(); // value -> first row index (1-based)

    rows.forEach((row, i) => {
      const rowNum = i + 1;
      const value = row['id_empleado'] ?? '';

      if (value === '') {
        issues.push({
          row: rowNum,
          column: 'id_empleado',
          value: '',
          message: 'Field is empty — id_empleado is required',
        });
        return; // skip duplicate check for empty values
      }

      if (seen.has(value)) {
        issues.push({
          row: rowNum,
          column: 'id_empleado',
          value,
          message: `Duplicate id_empleado — first seen at row ${seen.get(value)}`,
        });
      } else {
        seen.set(value, rowNum);
      }
    });

    return issues;
  },
};

// ---------------------------------------------------------------------------
// Rule: e_mail — required and must end with @arandadeduero.es
// ---------------------------------------------------------------------------

const ruleEmail = {
  id: 'email-required-domain',
  description: `e_mail must be filled and must end with ${REQUIRED_DOMAIN}`,
  validate(rows) {
    const issues = [];

    rows.forEach((row, i) => {
      const rowNum = i + 1;
      const value = row['e_mail'] ?? '';

      if (value === '') {
        issues.push({
          row: rowNum,
          column: 'e_mail',
          value: '',
          message: 'Field is empty — e_mail is required',
        });
        return;
      }

      if (!value.toLowerCase().endsWith(REQUIRED_DOMAIN)) {
        issues.push({
          row: rowNum,
          column: 'e_mail',
          value,
          message: `Email does not end with ${REQUIRED_DOMAIN}`,
        });
      }
    });

    return issues;
  },
};

// ---------------------------------------------------------------------------
// Rule: e_mail — must be unique across all rows
// ---------------------------------------------------------------------------

const ruleEmailUnique = {
  id: 'email-unique',
  description: 'e_mail must be unique — no two employees can share the same address',
  validate(rows) {
    const issues = [];
    const seen = new Map(); // email (lowercase) -> first row number

    rows.forEach((row, i) => {
      const rowNum = i + 1;
      const value = (row['e_mail'] ?? '').trim();

      if (value === '') return; // empty already caught by email-required-domain

      const key = value.toLowerCase();
      if (seen.has(key)) {
        issues.push({
          row: rowNum,
          column: 'e_mail',
          value,
          message: `Duplicate email — first seen at row ${seen.get(key)}`,
        });
      } else {
        seen.set(key, rowNum);
      }
    });

    return issues;
  },
};

// ---------------------------------------------------------------------------
// Rule: ID Responsable — if filled, must reference a known id_empleado
// ---------------------------------------------------------------------------

const ruleIdResponsable = {
  id: 'id_responsable-exists',
  description:
    'ID Responsable must reference an existing id_empleado (can be empty for the top-level director)',
  validate(rows) {
    const issues = [];

    // Build a set of all known id_empleado values (skip empty ones — already caught above)
    const knownIds = new Set(
      rows.map((r) => r['id_empleado'] ?? '').filter(Boolean),
    );

    rows.forEach((row, i) => {
      const rowNum = i + 1;
      const value = row['ID responsable'] ?? '';

      if (value === '') {
        // Empty is allowed (top-level director)
        return;
      }

      if (!knownIds.has(value)) {
        issues.push({
          row: rowNum,
          column: 'ID responsable',
          value,
          message: `ID responsable "${value}" does not match any id_empleado in the file`,
        });
      }
    });
    return issues;
  },
};

// ---------------------------------------------------------------------------
// Rule: Fecha de Baja — if filled, must be dd/mm/yyyy (2-digit day, month, 4-digit year)
// ---------------------------------------------------------------------------

// Matches exactly dd/mm/yyyy where dd and mm are 2 digits and yyyy is 4 digits.
const DATE_RE = /^\d{2}\/\d{2}\/\d{4}$/;

const ruleFechaDeBaja = {
  id: 'fecha-de-baja-format',
  description: 'Fecha de Baja, if filled, must be in dd/mm/yyyy format (2-digit day and month)',
  validate(rows) {
    const issues = [];

    rows.forEach((row, i) => {
      const rowNum = i + 1;
      const value = row['Fecha de Baja'] ?? '';

      if (value === '') return; // empty is fine — not all employees have a leaving date

      // Any date is accepted, including a future one: a leaving date can be
      // recorded in advance and only takes effect the day after (see sync-check).

      if (!DATE_RE.test(value)) {
        issues.push({
          row: rowNum,
          column: 'Fecha de Baja',
          value,
          message: `Expected dd/mm/yyyy format (e.g. 31/12/2025), got "${value}"`,
        });
      }
    });

    return issues;
  },
};

// ---------------------------------------------------------------------------
// Exported rule list — add new rules here
// ---------------------------------------------------------------------------

export const RULES = [ruleIdEmpleado, ruleEmail, ruleEmailUnique, ruleIdResponsable, ruleFechaDeBaja];
