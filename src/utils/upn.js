/**
 * Ensure a user input has the tenant domain appended if it lacks one.
 * If the value already contains '@', it's returned as-is.
 *
 * @param {string} value  - Raw user input
 * @param {string} domain - Tenant domain (e.g. from constants.js)
 * @returns {string}
 */
export function normalizeUpn(value, domain) {
  if (!value || !domain) return value;
  const v = value.trim();
  if (!v) return v;
  if (v.includes('@')) return v;
  return `${v}@${domain}`;
}

