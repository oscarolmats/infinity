/**
 * EPD CSV Parser
 * 
 * This module provides functions to parse EPD (Environmental Product Declaration) CSV files
 * and extract relevant climate impact data.
 */

class EpdParser {
  constructor() {
    this.debug = false;
  }

  /**
   * Enable or disable debug logging
   * @param {boolean} enabled - Whether to enable debug logging
   */
  setDebug(enabled) {
    this.debug = enabled;
  }

  /**
   * Log debug message if debug is enabled
   * @param {string} message - Debug message to log
   */
  log(message) {
    if (this.debug) {
      console.log(`🌱 [EpdParser] ${message}`);
    }
  }

  /**
   * Parse EPD CSV data and extract climate impact information
   * @param {string} csvText - Raw CSV text content
   * @returns {Object} Parsed EPD data with climate impact values
   */
  parseEpdCsv(csvText) {
    const lines = csvText.trim().split('\n');
    const headers = lines[0].split(';');
    const data = {};
    
    // Find relevant columns
    const nameIndex = headers.findIndex(h => h.includes('Name (en)'));
    const refQuantity = headers.findIndex(h => h.includes('Ref. quantity'));
    const refUnitIndex = headers.findIndex(h => h.includes('Ref. unit'));
    const moduleIndex = headers.findIndex(h => h.includes('Module'));
    const gwpIndex = headers.findIndex(h => h.includes('GWPtotal (A2)'));
    const urlIndex = headers.findIndex(h => h.includes('URL'));
    const declarationOwnerIndex = headers.findIndex(h => h.includes('Declaration owner'));
    const publicationDateIndex = headers.findIndex(h => h.includes('Publication date'));
    const registrationNumberIndex = headers.findIndex(h => h.includes('Registration number'));
    
    this.log('Headers found: ' + JSON.stringify({
      nameIndex,
      refQuantity,
      refUnitIndex, 
      moduleIndex,
      gwpIndex,
      urlIndex,
      declarationOwnerIndex,
      publicationDateIndex,
      registrationNumberIndex,
      headers: headers.slice(0, 10) // Show first 10 headers
    }));
    
    // Parse each module (A1-A3, A4, A5)
    for(let i = 1; i < lines.length; i++) {
      const values = lines[i].split(';');
      const module = values[moduleIndex];
      const gwp = parseFloat(values[gwpIndex]) || 0;
      
      this.log(`Row ${i}: module="${module}", gwp=${gwp}`);
      
      if(module === 'A1-A3') {
        data.name = values[nameIndex] || 'Unknown EPD';
        data.refQuantity = parseFloat(values[refQuantity]) || 1;
        data.refUnit = values[refUnitIndex] || 'kg';
        data.url = values[urlIndex] || '';
        data.declarationOwner = values[declarationOwnerIndex] || '';
        data.publicationDate = values[publicationDateIndex] || '';
        data.registrationNumber = values[registrationNumberIndex] || '';
        data.a1a3 = gwp;
        this.log(`A1-A3: name="${data.name}", unit="${data.refUnit}", url="${data.url}", owner="${data.declarationOwner}", date="${data.publicationDate}", reg="${data.registrationNumber}", gwp=${gwp}`);
      } else if(module === 'A4') {
        data.a4 = gwp;
        this.log(`A4: gwp=${gwp}`);
      } else if(module === 'A5') {
        data.a5 = gwp;
        this.log(`A5: gwp=${gwp}`);
      }
    }
    
    this.log('Final parsed data: ' + JSON.stringify(data));
    return data;
  }

  /**
   * Validate EPD data structure
   * @param {Object} epdData - Parsed EPD data to validate
   * @returns {Object} Validation result with isValid flag and errors array
   */
  validateEpdData(epdData) {
    const errors = [];
    
    if (!epdData.name || epdData.name === 'Unknown EPD') {
      errors.push('Missing or invalid product name');
    }
    
    if (!epdData.refUnit) {
      errors.push('Missing reference unit');
    }
    
    if (typeof epdData.a1a3 !== 'number' || epdData.a1a3 < 0) {
      errors.push('Invalid A1-A3 climate impact value');
    }
    
    if (typeof epdData.a4 !== 'number' || epdData.a4 < 0) {
      errors.push('Invalid A4 climate impact value');
    }
    
    if (typeof epdData.a5 !== 'number' || epdData.a5 < 0) {
      errors.push('Invalid A5 climate impact value');
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * Convert EPD data to display format
   * @param {Object} epdData - Parsed EPD data
   * @returns {Object} Formatted data for display
   */
  formatForDisplay(epdData) {
    const displayUnit = epdData.refUnit === 'qm' ? 'm²' : epdData.refUnit;

    return {
      name: epdData.name,
      referenceQuantity: epdData.refQuantity,
      unit: displayUnit,
      a1a3: epdData.a1a3 ? epdData.a1a3.toFixed(3) : '0',
      a4: epdData.a4 ? epdData.a4.toFixed(3) : '0',
      a5: epdData.a5 ? epdData.a5.toFixed(3) : '0',
      url: epdData.url,
      declarationOwner: epdData.declarationOwner,
      publicationDate: epdData.publicationDate,
      registrationNumber: epdData.registrationNumber
    };
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = EpdParser;
} else if (typeof window !== 'undefined') {
  window.EpdParser = EpdParser;
}
