/**
 * Central in-process cache for Microsoft Graph read results.
 *
 * All data here is process-lifetime: it lives as long as the CLI process runs
 * (typically one command invocation). Nothing is persisted to disk.
 *
 * Design principles:
 *  - Cache keys are always strings (UPN, object ID, or a normalised query key).
 *  - Cache entries are never evicted during a single process run — Graph data
 *    does not change mid-command in normal usage.
 *  - Write operations (updateUser, setManager, …) must call the appropriate
 *    invalidate* helper so stale entries are not served after a mutation.
 *  - The cache is intentionally simple: no TTL, no LRU, no external deps.
 */

// ---------------------------------------------------------------------------
// Auto-assigned groups — filtered out from all display/audit output.
// These are tenant-wide groups that every user belongs to automatically and
// carry no meaningful organisational information.
// ---------------------------------------------------------------------------

export const AUTO_GROUPS = new Set([
  'Todo Ayuntamiento',
  'Todos los usuarios',
  'Expertos 365',
]);

/** Returns groups with auto-assigned noise groups removed. */
export function filterAutoGroups(groups) {
  return groups.filter((g) => !AUTO_GROUPS.has(g.displayName ?? g));
}

// ---------------------------------------------------------------------------
// Storage maps
// ---------------------------------------------------------------------------

/** Map<userId|upn, userObject>  — full user objects from getUser() */
const _users = new Map();

/** Map<userId, managerObject|null>  — manager objects from getManager() */
const _managers = new Map();

/** Map<userId, Array<{id,displayName}>>  — group memberships from getUserGroups() */
const _userGroups = new Map();

/** Array<{id,displayName}> | null  — full group list from listAllGroups() */
let _allGroups = null;

/** Array<string> | null  — all department names from listAllDepartments() */
let _allDepartments = null;

/** Array<string> | null  — all job titles from listAllJobTitles() */
let _allJobTitles = null;

/** Map<groupId, displayName>  — group name lookup, populated from any group result */
const _groupNames = new Map();

// ---------------------------------------------------------------------------
// User cache
// ---------------------------------------------------------------------------

export function getCachedUser(key) {
  return _users.get(key) ?? null;
}

export function setCachedUser(user) {
  if (!user) return;
  // Store under both object ID and UPN so either key hits the cache
  if (user.id)                  _users.set(user.id, user);
  if (user.userPrincipalName)   _users.set(user.userPrincipalName, user);
}

/** Call after a successful updateUser() so the stale object is not served. */
export function invalidateUser(key) {
  const user = _users.get(key);
  if (user) {
    _users.delete(user.id);
    if (user.userPrincipalName) _users.delete(user.userPrincipalName);
  } else {
    _users.delete(key);
  }
}

// ---------------------------------------------------------------------------
// Manager cache
// ---------------------------------------------------------------------------

export function getCachedManager(userId) {
  // Use a sentinel to distinguish "cached as null" from "not cached"
  if (_managers.has(userId)) return { hit: true, value: _managers.get(userId) };
  return { hit: false, value: null };
}

export function setCachedManager(userId, manager) {
  _managers.set(userId, manager ?? null);
}

/** Call after a successful setManager() for this user. */
export function invalidateManager(userId) {
  _managers.delete(userId);
}

// ---------------------------------------------------------------------------
// User-groups cache
// ---------------------------------------------------------------------------

export function getCachedUserGroups(userId) {
  return _userGroups.has(userId) ? _userGroups.get(userId) : null;
}

export function setCachedUserGroups(userId, groups) {
  _userGroups.set(userId, groups ?? []);
  // Also populate the group-name lookup
  for (const g of groups ?? []) {
    if (g.id && g.displayName) _groupNames.set(g.id, g.displayName);
  }
}

/** Call after adding/removing the user from groups. */
export function invalidateUserGroups(userId) {
  _userGroups.delete(userId);
}

// ---------------------------------------------------------------------------
// All-groups list cache
// ---------------------------------------------------------------------------

export function getCachedAllGroups() {
  return _allGroups;
}

export function setCachedAllGroups(groups) {
  _allGroups = groups ?? [];
  for (const g of _allGroups) {
    if (g.id && g.displayName) _groupNames.set(g.id, g.displayName);
  }
}

// ---------------------------------------------------------------------------
// Group name lookup (populated as a side-effect of any group fetch)
// ---------------------------------------------------------------------------

export function getGroupName(groupId) {
  return _groupNames.get(groupId) ?? null;
}

export function setGroupName(groupId, displayName) {
  _groupNames.set(groupId, displayName);
}

// ---------------------------------------------------------------------------
// Departments cache
// ---------------------------------------------------------------------------

export function getCachedDepartments() {
  return _allDepartments;
}

export function setCachedDepartments(depts) {
  _allDepartments = depts ?? [];
}

// ---------------------------------------------------------------------------
// Job titles cache
// ---------------------------------------------------------------------------

export function getCachedJobTitles() {
  return _allJobTitles;
}

export function setCachedJobTitles(titles) {
  _allJobTitles = titles ?? [];
}

// ---------------------------------------------------------------------------
// Bulk-population helpers (used after fetchGroupsMap / listAllUsers results)
// ---------------------------------------------------------------------------

/**
 * Seed the user-groups cache from a Map<userId, groups[]> returned by fetchGroupsMap.
 * Called by list.js and validate-users.js after their batch fetch.
 */
export function seedGroupsMap(groupsMap) {
  for (const [userId, groups] of groupsMap) {
    setCachedUserGroups(userId, groups);
  }
}

/**
 * Seed the user cache from an array of user objects returned by listAllUsers.
 * Avoids redundant getUser() calls in subsequent operations within the same run.
 */
export function seedUsers(users) {
  for (const u of users) setCachedUser(u);
}

// ---------------------------------------------------------------------------
// Full reset (useful for tests)
// ---------------------------------------------------------------------------

export function _resetAllCaches() {
  _users.clear();
  _managers.clear();
  _userGroups.clear();
  _groupNames.clear();
  _allGroups       = null;
  _allDepartments  = null;
  _allJobTitles    = null;
}
