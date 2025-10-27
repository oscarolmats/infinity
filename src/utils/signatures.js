/**
 * Signature and row data utilities
 * Functions for creating unique signatures for rows and extracting row data
 */

/**
 * Create a unique signature for a row based on its data
 * @param {Array} rowData - Array of cell values
 * @param {string|null} layerChildOf - Layer hierarchy key
 * @returns {string} Unique signature
 */
export function getRowSignature(rowData, layerChildOf = null) {
  const baseSignature = JSON.stringify(rowData);
  // Include layer hierarchy in signature to differentiate nested layers
  return layerChildOf ? `${baseSignature}::layer::${layerChildOf}` : baseSignature;
}

/**
 * Extract row data from a TR element, excluding action and climate columns
 * @param {HTMLTableRowElement} tr - Table row element
 * @returns {Array|null} Array of cell values or null if table not found
 */
export function getRowDataFromTr(tr) {
  const table = tr.closest('table');
  if (!table) return null;

  const headers = Array.from(table.querySelectorAll('thead tr:first-child th')).map(th => th.textContent);
  const climateColIndex = headers.findIndex(h => h === 'Klimatresurs');
  const factorColIndex = headers.findIndex(h => h === 'Omräkningsfaktor');
  const unitColIndex = headers.findIndex(h => h === 'Omräkningsfaktor enhet');
  const wasteColIndex = headers.findIndex(h => h === 'Spillfaktor');
  const A1_A3ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A1-A3');
  const A4ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A4');
  const A5ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A5');
  const inbyggdViktColIndex = headers.findIndex(h => h === 'Inbyggd vikt');
  const inkoptViktColIndex = headers.findIndex(h => h === 'Inköpt vikt');
  const klimatA1A3ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A1-A3');
  const klimatA4ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A4');
  const klimatA5ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A5');

  const cells = Array.from(tr.children);
  const rowData = [];

  // Skip first cell (action column) and climate/factor/unit/waste columns if they exist
  for (let i = 1; i < cells.length; i++) {
    if ((climateColIndex !== -1 && i === climateColIndex) ||
        (factorColIndex !== -1 && i === factorColIndex) ||
        (unitColIndex !== -1 && i === unitColIndex) ||
        (wasteColIndex !== -1 && i === wasteColIndex) ||
        (A1_A3ColIndex !== -1 && i === A1_A3ColIndex) ||
        (A4ColIndex !== -1 && i === A4ColIndex) ||
        (A5ColIndex !== -1 && i === A5ColIndex) ||
        (inbyggdViktColIndex !== -1 && i === inbyggdViktColIndex) ||
        (inkoptViktColIndex !== -1 && i === inkoptViktColIndex) ||
        (klimatA1A3ColIndex !== -1 && i === klimatA1A3ColIndex) ||
        (klimatA4ColIndex !== -1 && i === klimatA4ColIndex) ||
        (klimatA5ColIndex !== -1 && i === klimatA5ColIndex)) {
      continue; // Skip climate, factor, unit, waste, emission, weight and climate impact columns
    }

    // Clone the cell and remove badges and decorations
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = cells[i].innerHTML;

    // Remove badges
    tempDiv.querySelectorAll('.badge-new').forEach(b => b.remove());
    // Remove toggle icons
    tempDiv.querySelectorAll('.group-toggle').forEach(t => t.remove());

    // Get clean text and remove layer indicators
    let text = tempDiv.textContent.trim();
    text = text.replace(/^\[Skikt \d+\/\d+\]\s*/, ''); // Remove "Skikt X/Y" prefix
    text = text.replace(/\s*\[\d+\s+skikt\]\s*$/, ''); // Remove "[N skikt]" suffix
    text = text.replace(/\s*\(\d+\)\s*$/, ''); // Remove "(N)" count suffix from groups

    rowData.push(text.trim());
  }

  return rowData;
}
