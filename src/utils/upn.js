/**
 * Ensure a user input has the tenant domain appended if it lacks one.
 * If the value already contains '@', it's returned as-is.
 *
 * @param {string} value  - Raw user input
 * @param {string} domain - Tenant domain, e.g. "arandadeduero.es"
 * @returns {string}
 */
export function normalizeUpn(value, domain) {
  if (!value || !domain) return value;
  const v = value.trim();
  if (!v) return v;
  if (v.includes('@')) return v;
  return `${v}@${domain}`;
}

/**
 * Extract the domain from a tenant UPN or from a domain string directly.
 * Accepts either "contoso.onmicrosoft.com", "contoso.es", or a full UPN.
 *
 * @param {string} domainOrUpn
 * @returns {string}
 */
export function parseDomain(domainOrUpn) {
  if (!domainOrUpn) return '';
  if (domainOrUpn.includes('@')) return domainOrUpn.split('@')[1];
  return domainOrUpn;
}
