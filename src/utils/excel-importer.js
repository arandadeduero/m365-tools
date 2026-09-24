import { readExcel } from './excel.js';

// Required fields for creating a new user
const REQUIRED_FIELDS = ['userPrincipalName', 'displayName', 'mailNickname'];

function toTitleCase(str) {
  if (!str) return '';
  return str.toLowerCase().split(' ').filter(Boolean).map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
}

/**
 * Maps Spanish Excel headers to internal property names.
 */
function normalizeRow(row, domain) {
  const normalized = {};
  const map = {
    'id_empleado': 'employeeId',
    'Departamento': 'department',
    'Trabajador': 'displayName',
    'Descripción': 'jobTitle',
    'descripción': 'jobTitle',
    'Fecha de Alta': 'employeeHireDate',
    'e_mail': 'userPrincipalName',
    'ID responsable': 'manager',
  };

  // Create a case-insensitive map for easier lookup
  const normalizedMap = Object.entries(map).reduce((acc, [k, v]) => {
    acc[k.toLowerCase().trim()] = v;
    return acc;
  }, {});

  for (const [key, value] of Object.entries(row)) {
    const lookupKey = key.toLowerCase().trim();
    const normalizedKey = normalizedMap[lookupKey] || key;
    normalized[normalizedKey] = value;
  }

  // Normalize Name: "SURNAME, NAME" -> "Name Surname"
  if (normalized.displayName) {
    let nameRaw = normalized.displayName.trim();
    if (nameRaw.includes(',')) {
      const parts = nameRaw.split(',');
      const surname = parts[0].trim();
      const firstName = parts.length > 1 ? parts[1].trim() : '';
      if (firstName) {
        normalized.givenName = toTitleCase(firstName);
        normalized.surname = toTitleCase(surname);
        normalized.displayName = `${normalized.givenName} ${normalized.surname}`;
      } else {
        normalized.displayName = toTitleCase(surname);
      }
    } else {
      normalized.displayName = toTitleCase(nameRaw);
    }
  }
  
  // Ensure UPN ends with domain
  if (normalized.userPrincipalName) {
    let upn = normalized.userPrincipalName.trim();
    if (!upn.includes('@')) {
      upn = `${upn}@${domain}`;
    }
    normalized.userPrincipalName = upn;
  }

  // Ensure required fields exist in a way rowToUserPayload expects
  if (!normalized.mailNickname && normalized.userPrincipalName) {
    normalized.mailNickname = normalized.userPrincipalName.split('@')[0];
  }
  
  return normalized;
}

/**
 * Parse an Excel file and return array of validated row objects.
 * @param {string} filePath - Path to the Excel file
 * @param {string} domain - Domain to enforce for UPNs
 * @returns {Promise<{ rows: Array, errors: Array }>} Parsed rows and any validation errors
 */
export async function parseExcel(filePath, domain) {
  const { rows } = await readExcel(filePath);

  const validatedRows = [];
  const errors = [];

  for (const [index, record] of rows.entries()) {
    const row = normalizeRow(record, domain);
    const rowNum = index + 2; // +2 because row 1 is header
    const rowErrors = [];

    // Validate required fields
    for (const field of REQUIRED_FIELDS) {
      if (!row[field] || row[field].trim() === '') {
        rowErrors.push(`Missing required field: ${field}`);
      }
    }

    // Validate UPN domain
    if (domain && row.userPrincipalName && !row.userPrincipalName.endsWith(`@${domain}`)) {
      rowErrors.push(`UPN must end with @${domain}`);
    }

    if (rowErrors.length > 0) {
      errors.push({ row: rowNum, upn: row.userPrincipalName || '(unknown)', errors: rowErrors });
      continue;
    }

    validatedRows.push(row);
  }

  return { rows: validatedRows, errors };
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
