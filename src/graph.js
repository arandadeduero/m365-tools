import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './auth.js';

/**
 * Build an authenticated Graph API client using the cached token.
 */
function buildClient() {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getAccessToken();
        done(null, token);
      } catch (err) {
        done(err, null);
      }
    },
  });
}

// ---------------------------------------------------------------------------
// User operations
// ---------------------------------------------------------------------------

/**
 * Create a new user in Azure AD.
 * @param {object} userData - User properties matching Graph API user resource
 * @returns {object} Created user object
 */
export async function createUser(userData) {
  const client = buildClient();
  return client.api('/users').post(userData);
}

/**
 * Get a user by UPN or object ID.
 * @param {string} identifier - UPN (email) or object ID
 * @returns {object|null} User object or null if not found
 */
export async function getUser(identifier) {
  const client = buildClient();
  try {
    return await client
      .api(`/users/${encodeURIComponent(identifier)}`)
      .select([
        'id', 'userPrincipalName', 'displayName', 'givenName', 'surname',
        'mail', 'mailNickname', 'jobTitle', 'department', 'officeLocation',
        'mobilePhone', 'businessPhones', 'usageLocation', 'accountEnabled',
        'city', 'country', 'postalCode', 'state', 'streetAddress',
        'companyName', 'employeeId', 'employeeType', 'employeeHireDate',
        'preferredLanguage', 'userType',
      ])
      .get();
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

/**
 * Update an existing user.
 * @param {string} identifier - UPN or object ID
 * @param {object} updates - Properties to update
 */
export async function updateUser(identifier, updates) {
  const client = buildClient();
  return client.api(`/users/${encodeURIComponent(identifier)}`).patch(updates);
}

/**
 * Search users by display name or mail (prefix search).
 * @param {string} query - Search string
 * @param {number} limit - Max results (default 25)
 * @returns {Array} Array of user objects
 */
export async function searchUsers(query, limit = 25) {
  const client = buildClient();

  // Try $search first (requires ConsistencyLevel header), fallback to $filter
  try {
    const result = await client
      .api('/users')
      .header('ConsistencyLevel', 'eventual')
      .search(`"displayName:${query}" OR "mail:${query}" OR "userPrincipalName:${query}"`)
      .select(['id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail', 'jobTitle', 'department', 'accountEnabled'])
      .top(limit)
      .get();
    return result.value || [];
  } catch {
    // Fallback: filter by displayName or mail startsWith
    const safe = query.replace(/'/g, "''");  // OData single-quote escaping
    const filter = `startsWith(displayName,'${safe}') or startsWith(mail,'${safe}') or startsWith(userPrincipalName,'${safe}')`;
    const result = await client
      .api('/users')
      .filter(filter)
      .select(['id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail', 'jobTitle', 'department', 'accountEnabled'])
      .top(limit)
      .get();
    return result.value || [];
  }
}

/**
 * Set or update the manager for a user.
 * @param {string} userId - UPN or object ID of the user
 * @param {string} managerUpn - UPN of the manager
 */
export async function setManager(userId, managerUpn) {
  // First resolve manager to their object ID
  const manager = await getUser(managerUpn);
  if (!manager) {
    throw new Error(`Manager not found: ${managerUpn}`);
  }

  const client = buildClient();
  await client.api(`/users/${encodeURIComponent(userId)}/manager/$ref`).put({
    '@odata.id': `https://graph.microsoft.com/v1.0/users/${manager.id}`,
  });
}

/**
 * Get the current manager of a user.
 * @param {string} userId - UPN or object ID
 * @returns {object|null} Manager user object or null
 */
export async function getManager(userId) {
  const client = buildClient();
  try {
    return await client
      .api(`/users/${encodeURIComponent(userId)}/manager`)
      .select(['id', 'displayName', 'userPrincipalName', 'jobTitle'])
      .get();
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

/**
 * List all users in the tenant with pagination.
 *
 * @param {object} opts
 * @param {boolean} opts.onlyDisabled   - List only disabled users (accountEnabled=false)
 * @param {boolean} opts.checkManager   - Fetch manager UPN for each user (parallel batches)
 * @param {Function} opts.onProgress    - Called with (fetched) during pagination/manager fetch
 * @returns {Array} Users with optional managerUpn string
 */
export async function listAllUsers({ onlyDisabled = false, checkManager = false, onProgress } = {}) {
  const client = buildClient();
  const select = [
    'id', 'userPrincipalName', 'displayName', 'givenName', 'surname',
    'mail', 'jobTitle', 'department', 'companyName', 'officeLocation',
    'accountEnabled', 'userType',
  ];

  // Filter server-side: active/disabled + exclude external guest users (#EXT#)
  const accountFilter = onlyDisabled
    ? `accountEnabled eq false`
    : `accountEnabled eq true`;

  let nextLink = null;
  const users = [];

  // First page uses filter; subsequent pages use nextLink directly
  do {
    let req;
    if (nextLink) {
      req = client.api(nextLink);
    } else {
      req = client.api('/users').filter(accountFilter).select(select).top(999);
    }
    const result = await req.get();
    // Exclude external/guest users: their UPN contains #EXT#
    const page = (result.value || []).filter(
      (u) => !u.userPrincipalName || !u.userPrincipalName.includes('#EXT#')
    );
    users.push(...page);
    if (onProgress) onProgress(users.length);
    nextLink = result['@odata.nextLink'] || null;
  } while (nextLink);

  if (!checkManager) return users;

  // Fetch manager UPN for all users in parallel batches of 20
  const BATCH = 20;
  for (let i = 0; i < users.length; i += BATCH) {
    const slice = users.slice(i, i + BATCH);
    const results = await Promise.allSettled(
      slice.map((u) =>
        client
          .api(`/users/${u.id}/manager`)
          .select(['userPrincipalName', 'displayName'])
          .get()
          .then((m) => ({ upn: m.userPrincipalName, name: m.displayName }))
          .catch(() => null)
      )
    );
    for (const [idx, r] of results.entries()) {
      slice[idx].manager = r.status === 'fulfilled' ? r.value : null;
    }
    if (onProgress) onProgress(i + BATCH);
  }

  return users;
}

/**
 * Check if a user exists by UPN.
 * @param {string} upn - User Principal Name
 * @returns {object|null} User with id and displayName, or null
 */
export async function findUserByUpn(upn) {
  const client = buildClient();
  try {
    const result = await client
      .api('/users')
      .filter(`userPrincipalName eq '${upn}'`)
      .select(['id', 'userPrincipalName', 'displayName'])
      .get();
    return result.value && result.value.length > 0 ? result.value[0] : null;
  } catch {
    return null;
  }
}

/**
 * Get all distinct non-empty department values across all active users.
 * @returns {string[]} Sorted unique departments
 */
export async function listAllDepartments() {
  const client = buildClient();
  const depts = new Set();

  let nextLink = null;
  do {
    let req;
    if (nextLink) {
      req = client.api(nextLink);
    } else {
      req = client
        .api('/users')
        .filter('accountEnabled eq true')
        .select(['department', 'userPrincipalName'])
        .top(999);
    }
    const result = await req.get();
    for (const u of result.value || []) {
      if (u.userPrincipalName && u.userPrincipalName.includes('#EXT#')) continue;
      if (u.department && u.department.trim() !== '') {
        depts.add(u.department.trim());
      }
    }
    nextLink = result['@odata.nextLink'] || null;
  } while (nextLink);

  return [...depts].sort((a, b) => a.localeCompare(b));
}

/**
 * Get all active users belonging to a specific department.
 * @param {string} department - Exact department name
 * @returns {object[]} Array of user objects
 */
export async function getUsersByDepartment(department) {
  const client = buildClient();
  const select = ['id', 'userPrincipalName', 'displayName', 'jobTitle', 'department'];
  const filter = `accountEnabled eq true and department eq '${department.replace(/'/g, "''")}'`;

  const users = [];
  let nextLink = null;

  do {
    let req;
    if (nextLink) {
      req = client.api(nextLink);
    } else {
      req = client.api('/users').filter(filter).select(select).top(999);
    }
    const result = await req.get();
    const page = (result.value || []).filter(
      (u) => !u.userPrincipalName || !u.userPrincipalName.includes('#EXT#')
    );
    users.push(...page);
    nextLink = result['@odata.nextLink'] || null;
  } while (nextLink);

  return users;
}

/**
 * Get all distinct non-empty jobTitle values across all active users.
 * Paginates through the full user list.
 * @returns {string[]} Sorted unique job titles
 */
export async function listAllJobTitles() {
  const client = buildClient();
  const select = ['jobTitle'];
  const titles = new Set();

  let nextLink = null;
  do {
    let req;
    if (nextLink) {
      req = client.api(nextLink);
    } else {
      req = client
        .api('/users')
        .filter('accountEnabled eq true')
        .select(select)
        .top(999);
    }
    const result = await req.get();
    for (const u of result.value || []) {
      // Skip external/guest users (#EXT# in UPN)
      if (u.userPrincipalName && u.userPrincipalName.includes('#EXT#')) continue;
      if (u.jobTitle && u.jobTitle.trim() !== '') {
        titles.add(u.jobTitle.trim());
      }
    }
    nextLink = result['@odata.nextLink'] || null;
  } while (nextLink);

  return [...titles].sort((a, b) => a.localeCompare(b));
}

/**
 * Get all active users with a specific jobTitle.
 * @param {string} jobTitle - Exact job title to filter by
 * @returns {object[]} Array of user objects
 */
export async function getUsersByJobTitle(jobTitle) {
  const client = buildClient();
  const select = ['id', 'userPrincipalName', 'displayName', 'jobTitle', 'department'];

  // Graph API filter is case-insensitive for eq on jobTitle
  const filter = `accountEnabled eq true and jobTitle eq '${jobTitle.replace(/'/g, "''")}'`;

  const users = [];
  let nextLink = null;

  do {
    let req;
    if (nextLink) {
      req = client.api(nextLink);
    } else {
      req = client.api('/users').filter(filter).select(select).top(999);
    }
    const result = await req.get();
    const page = (result.value || []).filter(
      (u) => !u.userPrincipalName || !u.userPrincipalName.includes('#EXT#')
    );
    users.push(...page);
    nextLink = result['@odata.nextLink'] || null;
  } while (nextLink);

  return users;
}

/**
 * Fetch the manager ID for every user in the list using Graph $batch (20 requests per call).
 * Returns a Map<userId, managerId|null>.
 *
 * @param {Array}    users      - Array of user objects with at least { id }
 * @param {Function} onProgress - Optional progress callback (processed, total)
 * @returns {Map<string, string|null>}
 */
export async function fetchManagerMap(users, onProgress) {
  const token = await getAccessToken();
  const map   = new Map();

  const BATCH = 20;

  for (let i = 0; i < users.length; i += BATCH) {
    const slice = users.slice(i, i + BATCH);

    const requests = slice.map((u, idx) => ({
      id:     String(idx),
      method: 'GET',
      url:    `/users/${u.id}/manager?$select=id`,
    }));

    const resp = await fetch('https://graph.microsoft.com/v1.0/$batch', {
      method:  'POST',
      headers: {
        Authorization:  `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ requests }),
    });

    if (!resp.ok) {
      const errBody = await resp.json().catch(() => ({}));
      throw new Error(`Graph $batch failed (${resp.status}): ${errBody?.error?.message || resp.statusText}`);
    }

    const data = await resp.json();

    for (const res of data.responses || []) {
      const idx    = parseInt(res.id, 10);
      const userId = slice[idx]?.id;
      if (!userId) continue;

      if (res.status === 200 && res.body?.id) {
        map.set(userId, res.body.id);
      } else {
        map.set(userId, null);
      }
    }

    if (onProgress) onProgress(Math.min(i + BATCH, users.length), users.length);
  }

  return map;
}
