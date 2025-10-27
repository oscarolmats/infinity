/**
 * Global data store for application state
 * Manages layer data, climate data, and undo/redo stacks
 */

// Storage for layers and climate resources
export const layerData = new Map(); // key: row signature -> { count, thicknesses, layerKey, mixedLayerConfigs, etc. }
export const climateData = new Map(); // key: row signature or layerKey -> climate resource data

// Undo/Redo functionality
export const undoStack = [];
export const redoStack = [];
export const maxUndoSteps = 50; // Limit history to prevent memory issues
export let isRestoringState = false; // Flag to prevent saving during restore

/**
 * Set the restoring state flag
 * @param {boolean} value - Whether we're currently restoring state
 */
export function setRestoringState(value) {
  isRestoringState = value;
}

/**
 * Clear all data stores
 */
export function clearAllData() {
  layerData.clear();
  climateData.clear();
  undoStack.length = 0;
  redoStack.length = 0;
}
