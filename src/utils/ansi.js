/**
 * Remove ANSI escape codes from a string.
 * Covers SGR (color/style), cursor movement, and other CSI sequences
 * emitted by chalk and similar libraries.
 * Used when calculating display widths for terminal table rendering.
 * @param {string} str
 * @returns {string}
 */
export function stripAnsi(str) {
  // Matches all CSI sequences: ESC [ <params> <final-byte>
  return String(str).replace(/\x1B\[[\x20-\x3F]*[\x40-\x7E]/g, '');
}
