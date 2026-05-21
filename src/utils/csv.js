import { parse } from 'csv-parse/sync';
import { readFile } from 'node:fs/promises';

// Required fields for creating a new user
const REQUIRED_FIELDS = ['userPrincipalName', 'displayName', 'mailNickname', 'password'];

// All supported CSV columns and their mapping to Graph API properties
export const FIELD_MAP = {
  // Identity
  userPrincipalName: 'userPrincipalName',
  displayName: 'displayName',
  givenName: 'givenName',
  surname: 'surname',
  mailNickname: 'mailNickname',
  mail: 'mail',

  // Account
  accountEnabled: 'accountEnabled',
  password: null, // handled separately -> passwordProfile
  forceChangePasswordNextSignIn: null, // handled separately -> passwordProfile
  usageLocation: 'usageLocation',
  preferredLanguage: 'preferredLanguage',
  userType: 'userType',

  // Job info
  jobTitle: 'jobTitle',
  department: 'department',
  companyName: 'companyName',
  employeeId: 'employeeId',
  employeeType: 'employeeType',
  employeeHireDate: 'employeeHireDate',

  // Contact
  mobilePhone: 'mobilePhone',
  businessPhone: null, // -> businessPhones[]
  officeLocation: 'officeLocation',

  // Address
  streetAddress: 'streetAddress',
  city: 'city',
  state: 'state',
  postalCode: 'postalCode',
  country: 'country',

  // Org chart
  manager: null, // handled separately after user creation/update
};

/**
 * Parse a CSV file and return array of validated row objects.
 * @param {string} filePath - Path to the CSV file
 * @returns {Promise<{ rows: Array, errors: Array }>} Parsed rows and any validation errors
 */
export async function parseCSV(filePath) {
  const content = await readFile(filePath, 'utf8');

  const records = parse(content, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
    bom: true, // Handle BOM from Excel-exported CSVs
  });

  const rows = [];
  const errors = [];

  for (const [index, record] of records.entries()) {
    const rowNum = index + 2; // +2 because row 1 is header
    const rowErrors = [];

    // Validate required fields
    for (const field of REQUIRED_FIELDS) {
      if (!record[field] || record[field].trim() === '') {
        rowErrors.push(`Missing required field: ${field}`);
      }
    }

    if (rowErrors.length > 0) {
      errors.push({ row: rowNum, upn: record.userPrincipalName || '(unknown)', errors: rowErrors });
      continue;
    }

    rows.push(record);
  }

  return { rows, errors };
}

/**
 * Convert a CSV row into a Graph API user payload.
 * @param {object} row - CSV row object
 * @returns {{ userPayload: object, manager: string|null }} User data and optional manager UPN
 */
export function rowToUserPayload(row) {
  const userPayload = {};

  // Scalar string fields
  const scalarFields = [
    'userPrincipalName', 'displayName', 'givenName', 'surname', 'mailNickname',
    'mail', 'jobTitle', 'department', 'companyName', 'employeeId', 'employeeType',
    'officeLocation', 'mobilePhone', 'streetAddress', 'city', 'state',
    'postalCode', 'country', 'usageLocation', 'preferredLanguage', 'userType',
  ];

  for (const field of scalarFields) {
    if (row[field] && row[field].trim() !== '') {
      userPayload[field] = row[field].trim();
    }
  }

  // Boolean: accountEnabled
  if (row.accountEnabled !== undefined && row.accountEnabled !== '') {
    userPayload.accountEnabled = row.accountEnabled.toLowerCase() !== 'false';
  } else {
    userPayload.accountEnabled = true; // default: enabled
  }

  // businessPhones array
  if (row.businessPhone && row.businessPhone.trim() !== '') {
    userPayload.businessPhones = [row.businessPhone.trim()];
  }

  // employeeHireDate: validate ISO 8601
  if (row.employeeHireDate && row.employeeHireDate.trim() !== '') {
    const date = new Date(row.employeeHireDate.trim());
    if (!isNaN(date.getTime())) {
      userPayload.employeeHireDate = date.toISOString();
    }
  }

  // passwordProfile (only for creation, or if password column provided)
  if (row.password && row.password.trim() !== '') {
    const forceChange = row.forceChangePasswordNextSignIn !== undefined
      ? row.forceChangePasswordNextSignIn.toLowerCase() !== 'false'
      : true;

    userPayload.passwordProfile = {
      password: row.password.trim(),
      forceChangePasswordNextSignIn: forceChange,
    };
  }

  // Extract manager (handled separately)
  const manager = row.manager && row.manager.trim() !== '' ? row.manager.trim() : null;

  return { userPayload, manager };
}

/**
 * Strip passwordProfile from an update payload (passwords not typically updated via CSV update).
 */
export function stripPasswordFromPayload(payload) {
  const { passwordProfile, ...rest } = payload;
  return rest;
}
