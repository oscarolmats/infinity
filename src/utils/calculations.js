/**
 * Calculation and parsing utilities
 * Functions for parsing and calculating numeric values
 */

/**
 * Parse a value into a number, handling various formats
 * Handles spaces, commas as decimal separator, and null/undefined
 * @param {any} v - Value to parse
 * @returns {number} Parsed number or NaN if invalid
 */
export function parseNumberLike(v) {
  if (v == null) return NaN;
  const s = String(v).trim().replace(/\s+/g, '').replace(',', '.');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}
