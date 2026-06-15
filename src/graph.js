import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './auth.js';
import {
  getCachedUser, setCachedUser, invalidateUser,
  getCachedManager, setCachedManager, invalidateManager,
  getCachedUserGroups, setCachedUserGroups, invalidateUserGroups,
  getCachedAllGroups, setCachedAllGroups,
  getCachedDepartments, setCachedDepartments,
  getCachedJobTitles, setCachedJobTitles,
  seedUsers, seedGroupsMap,
  setGroupName,
} from './utils/cache.js';

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
  const cached = getCachedUser(identifier);
  if (cached) return cached;

  const client = buildClient();
  try {
    const user = await client
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
    setCachedUser(user);
    return user;
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
  const result = await client.api(`/users/${encodeURIComponent(identifier)}`).patch(updates);
  invalidateUser(identifier);
  return result;
}

/**
 * Get all licences currently assigned to a user.
 * @param {string} identifier - UPN or object ID
 * @returns {Promise<Array<{skuId: string, skuPartNumber: string}>>}
 */
export async function getUserLicenses(identifier) {
  const client = buildClient();
  const result = await client
    .api(`/users/${encodeURIComponent(identifier)}/licenseDetails`)
    .select(['skuId', 'skuPartNumber'])
    .get();
  return result.value ?? [];
}

/**
 * Remove all licences from a user. No-op if the user has none.
 * @param {string} identifier   - UPN or object ID
 * @param {Array}  [licenses]   - Optional pre-fetched license list (skips extra API call)
 * @returns {Promise<void>}
 */
export async function removeAllLicenses(identifier, licenses) {
  const list = licenses ?? await getUserLicenses(identifier);
  if (list.length === 0) return;
  const client = buildClient();
  await client.api(`/users/${encodeURIComponent(identifier)}/assignLicense`).post({
    addLicenses: [],
    removeLicenses: list.map((l) => l.skuId),
  });
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
/**
 * Set or update the manager for a user.
 * @param {string} userId      - UPN or object ID of the user
 * @param {string} managerUpn  - UPN of the manager
 * @param {string} [managerId] - Optional pre-resolved object ID; skips getUser() lookup when provided
 */
export async function setManager(userId, managerUpn, managerId) {
  // Resolve manager object ID — use cache or the optional pre-resolved value first
  let resolvedId = managerId;
  if (!resolvedId) {
    const manager = await getUser(managerUpn);
    if (!manager) throw new Error(`Manager not found: ${managerUpn}`);
    resolvedId = manager.id;
  }

  const client = buildClient();
  await client.api(`/users/${encodeURIComponent(userId)}/manager/$ref`).put({
    '@odata.id': `https://graph.microsoft.com/v1.0/users/${resolvedId}`,
  });
  invalidateManager(userId);
}

/**
 * Get the current manager of a user.
 * @param {string} userId - UPN or object ID
 * @returns {object|null} Manager user object or null
 */
export async function getManager(userId) {
  const { hit, value } = getCachedManager(userId);
  if (hit) return value;

  const client = buildClient();
  try {
    const manager = await client
      .api(`/users/${encodeURIComponent(userId)}/manager`)
      .select(['id', 'displayName', 'userPrincipalName', 'jobTitle'])
      .get();
    setCachedManager(userId, manager);
    return manager;
  } catch (err) {
    if (err.statusCode === 404) {
      setCachedManager(userId, null);
      return null;
    }
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
    'accountEnabled', 'userType', 'employeeId', 'employeeType',
    'signInActivity', 'lastPasswordChangeDateTime',
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

  if (!checkManager) {
    seedUsers(users);
    return users;
  }

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

  seedUsers(users);
  return users;
}

/**
 * Check if a user exists by UPN.
 * @param {string} upn - User Principal Name
 * @returns {object|null} User with id and displayName, or null
 */
export async function findUserByUpn(upn) {
  const cached = getCachedUser(upn);
  if (cached) return cached;

  const client = buildClient();
  try {
    const result = await client
      .api('/users')
      .filter(`userPrincipalName eq '${upn}'`)
      .select(['id', 'userPrincipalName', 'displayName'])
      .get();
    const user = result.value && result.value.length > 0 ? result.value[0] : null;
    if (user) setCachedUser(user);
    return user;
  } catch {
    return null;
  }
}

/**
 * Get all distinct non-empty department values across all active users.
 * @returns {string[]} Sorted unique departments
 */
export async function listAllDepartments() {
  const cached = getCachedDepartments();
  if (cached) return cached;

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

  const sorted = [...depts].sort((a, b) => a.localeCompare(b));
  setCachedDepartments(sorted);
  return sorted;
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
  const cached = getCachedJobTitles();
  if (cached) return cached;

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

  const sortedTitles = [...titles].sort((a, b) => a.localeCompare(b));
  setCachedJobTitles(sortedTitles);
  return sortedTitles;
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

// ---------------------------------------------------------------------------
// Group membership
// ---------------------------------------------------------------------------

/**
 * Search groups by display name prefix.
 * @param {string} query
 * @returns {Promise<Array<{id: string, displayName: string}>>}
 */
export async function searchGroups(query) {
  const client = buildClient();
  const safe = query.replace(/'/g, "''");
  const result = await client
    .api('/groups')
    .filter(`startsWith(displayName,'${safe}')`)
    .select(['id', 'displayName'])
    .top(25)
    .get();
  return result.value ?? [];
}

/**
 * Add a user to a group.
 * Silently succeeds if the user is already a member (409 Conflict).
 * @param {string} groupId
 * @param {string} userId
 */
export async function addMemberToGroup(groupId, userId) {
  const client = buildClient();
  try {
    await client.api(`/groups/${groupId}/members/$ref`).post({
      '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
    });
  } catch (err) {
    // 409 = already a member — treat as success
    if (err.statusCode === 409) return;
    throw err;
  }
  invalidateUserGroups(userId);
}

/**
 * Remove a user from a group.
 * Silently succeeds if the user is not a member (404).
 * @param {string} groupId
 * @param {string} userId
 */
export async function removeMemberFromGroup(groupId, userId) {
  const client = buildClient();
  try {
    await client.api(`/groups/${groupId}/members/${userId}/$ref`).delete();
  } catch (err) {
    if (err.statusCode === 404) return;
    throw err;
  }
  invalidateUserGroups(userId);
}

/**
 * List all groups in the tenant (paginated).
 * @returns {Promise<Array<{id: string, displayName: string}>>}
 */
export async function listAllGroups() {
  const cached = getCachedAllGroups();
  if (cached) return cached;

  const client = buildClient();
  const groups = [];
  let url = '/groups?$select=id,displayName&$top=100&$orderby=displayName';

  while (url) {
    const page = await client.api(url).get();
    groups.push(...(page.value ?? []));
    url = page['@odata.nextLink'] ?? null;
  }

  const sorted = groups.sort((a, b) => a.displayName.localeCompare(b.displayName, 'es'));
  setCachedAllGroups(sorted);
  return sorted;
}

/**
 * Get all groups a user is a direct member of.
 * Returns only security groups and Microsoft 365 groups (not roles).
 * @param {string} identifier - UPN or object ID
 * @returns {Promise<Array<{id: string, displayName: string}>>}
 */
export async function getUserGroups(identifier) {
  // For cache lookup we need the user ID; if the caller passes a UPN we do a
  // best-effort lookup from the user cache first before falling back to the API.
  const cachedUser = getCachedUser(identifier);
  const userId = cachedUser?.id ?? identifier;

  const cached = getCachedUserGroups(userId);
  if (cached) return cached;

  const client = buildClient();
  const groups = [];
  let url = `/users/${encodeURIComponent(identifier)}/memberOf/microsoft.graph.group?$select=id,displayName&$top=100`;

  while (url) {
    const page = await client.api(url).get();
    groups.push(...(page.value ?? []));
    url = page['@odata.nextLink'] ?? null;
  }

  setCachedUserGroups(userId, groups);
  return groups;
}

/**
 * Fetch group memberships for a list of users using Graph $batch (20 per batch).
 * Returns a Map<userId, Array<{id, displayName}>>
 * @param {Array<{id: string}>} users
 * @param {function} [onProgress]
 */
export async function fetchGroupsMap(users, onProgress) {
  const token = await getAccessToken();
  const BATCH = 20;
  const map = new Map();

  for (let i = 0; i < users.length; i += BATCH) {
    const slice = users.slice(i, i + BATCH);

    const requests = slice.map((u, idx) => ({
      id: String(idx),
      method: 'GET',
      url: `/users/${encodeURIComponent(u.id)}/memberOf/microsoft.graph.group?$select=id,displayName&$top=100`,
    }));

    const resp = await fetch('https://graph.microsoft.com/v1.0/$batch', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ requests }),
    });

    if (!resp.ok) {
      const errBody = await resp.json().catch(() => ({}));
      throw new Error(`Graph $batch failed (${resp.status}): ${errBody?.error?.message || resp.statusText}`);
    }

    const data = await resp.json();

    for (const res of data.responses ?? []) {
      const idx = parseInt(res.id, 10);
      const userId = slice[idx]?.id;
      if (!userId) continue;

      if (res.status === 200) {
        map.set(userId, res.body?.value ?? []);
      } else {
        map.set(userId, []);
      }
    }

    if (onProgress) onProgress(Math.min(i + BATCH, users.length), users.length);
  }

  seedGroupsMap(map);
  return map;
}

// ---------------------------------------------------------------------------
// Mailbox settings
// ---------------------------------------------------------------------------

/**
 * Get the mailbox timezone settings for a user, including both the top-level
 * calendar timezone and the workingHours timezone (which Teams and Outlook use
 * for free/busy and availability display).
 *
 * Returns null if the user has no Exchange mailbox (no license, 404, 403, 400).
 *
 * @param {string} identifier - UPN or object ID
 * @returns {Promise<{
 *   timeZone: string|null,
 *   workingHoursTimeZone: string|null,
 *   workingHours: object|null,
 * }|null>}
 */
export async function getMailboxTzSettings(identifier) {
  const client = buildClient();
  try {
    // Do NOT use .select() here — $select is not supported on the mailboxSettings
    // singleton endpoint and causes a 400, silently returning null for every user.
    const result = await client
      .api(`/users/${encodeURIComponent(identifier)}/mailboxSettings`)
      .get();

    return {
      timeZone:             result?.timeZone ?? null,
      workingHoursTimeZone: result?.workingHours?.timeZone?.name ?? null,
      workingHours:         result?.workingHours ?? null,
      _raw:                 result,
      _error:               null,
    };
  } catch (err) {
    const status = err.statusCode ?? err.status ?? err.response?.status;
    const code   = err.code ?? '';

    const noMailbox =
      status === 404 ||
      String(code).includes('MailboxNotEnabledForRESTAPI') ||
      String(code).includes('ResourceNotFound') ||
      String(code).includes('ErrorMailboxStoreUnavailable');

    const accessDenied =
      status === 403 ||
      String(code).includes('ErrorAccessDenied');

    const badRequest = status === 400;

    if (noMailbox || accessDenied || badRequest) {
      return {
        _raw:    null,
        _error:  { status, code, message: err.message, body: err.body ?? null },
        // Expose the reason so fix-timezone.js can categorise correctly
        _reason: noMailbox ? 'no-mailbox' : accessDenied ? 'access-denied' : 'bad-request',
      };
    }
    throw err;
  }
}

/**
 * Set the mailbox timezone for a user, updating both the top-level calendar
 * timezone AND the workingHours timezone in a single PATCH call.
 *
 * The workingHours schedule (days of week, start/end times) is preserved
 * exactly as-is; only the timezone inside workingHours is changed.
 * If the user has no workingHours configured, only the top-level timeZone
 * is patched.
 *
 * @param {string}      identifier    - UPN or object ID
 * @param {string}      timeZone      - Windows timezone string (e.g. "Romance Standard Time")
 * @param {object|null} workingHours  - Existing workingHours object from getMailboxTzSettings
 * @returns {Promise<object>} Updated mailboxSettings fragment
 */
export async function setMailboxTzSettings(identifier, timeZone, workingHours) {
  const client = buildClient();

  const body = { timeZone };

  if (workingHours) {
    body.workingHours = {
      daysOfWeek: workingHours.daysOfWeek,
      startTime:  workingHours.startTime,
      endTime:    workingHours.endTime,
      timeZone:   { name: timeZone },
    };
  }

  return client
    .api(`/users/${encodeURIComponent(identifier)}/mailboxSettings`)
    .patch(body);
}
