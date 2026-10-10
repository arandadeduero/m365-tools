export const DOMAIN = 'arandadeduero.es';
export const REQUIRED_DOMAIN = `@${DOMAIN}`;

// Special accounts (shared mailboxes, external consultants, meter readers…) that
// are expected to exist in M365 without a row in the employee Excel file.
// sync-check does not count them as cloud-only issues, but lists them as a
// reminder to review them periodically. Usernames only, lowercase.
export const SPECIAL_ACCOUNTS = [
  'asac',
  'info_practicas1',
  'lector01agua',
  'consultorahyk',
  'lector02agua',
].map((name) => `${name}${REQUIRED_DOMAIN}`);
