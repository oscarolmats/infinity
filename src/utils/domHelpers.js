/**
 * DOM helper utilities
 * Functions for accessing and manipulating DOM elements
 */

/**
 * Get the main table element from the output container
 * @param {HTMLElement} outputElement - The output container element
 * @returns {HTMLTableElement|null} The table element or null if not found
 */
export function getTable(outputElement) {
  return outputElement.querySelector('table');
}
