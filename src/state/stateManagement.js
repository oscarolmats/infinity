/**
 * State management for undo/redo functionality
 * Factory functions that create state management functions with proper context
 */

import { undoStack, redoStack, maxUndoSteps, isRestoringState, setRestoringState } from './dataStore.js';

/**
 * Create state management functions with application context
 * @param {Object} context - Application context
 * @param {HTMLElement} context.output - Output container element
 * @param {Object} context.dataRefs - References to mutable data (layerData, climateData, lastRows, lastHeaders)
 * @param {HTMLSelectElement} context.groupBySelect - Group by select element
 * @param {Function} context.reattachTableEventListeners - Function to reattach event listeners
 * @param {Function} context.updateClimateSummary - Function to update climate summary
 * @param {HTMLButtonElement} context.undoBtn - Undo button element
 * @param {HTMLButtonElement} context.redoBtn - Redo button element
 * @returns {Object} State management functions
 */
export function createStateManagement(context) {
  const {
    output,
    dataRefs,
    groupBySelect,
    reattachTableEventListeners,
    updateClimateSummary,
    undoBtn,
    redoBtn
  } = context;

  /**
   * Update undo/redo button states
   */
  function updateUndoRedoButtons() {
    if (undoBtn) {
      undoBtn.disabled = undoStack.length === 0;
    }
    if (redoBtn) {
      redoBtn.disabled = redoStack.length === 0;
    }
  }

  /**
   * Restore state from a saved state object
   * @param {Object} state - Saved state
   */
  function restoreState(state) {
    if (!state) return;

    setRestoringState(true);

    output.innerHTML = state.outputHTML;
    dataRefs.layerData.clear();
    state.layerData.forEach((value, key) => dataRefs.layerData.set(key, value));
    dataRefs.climateData.clear();
    state.climateData.forEach((value, key) => dataRefs.climateData.set(key, value));
    dataRefs.lastRows = state.lastRows ? JSON.parse(JSON.stringify(state.lastRows)) : null;
    dataRefs.lastHeaders = state.lastHeaders ? [...state.lastHeaders] : null;

    if (groupBySelect && state.groupByValue !== undefined) {
      groupBySelect.value = state.groupByValue;
    }

    // Re-attach event listeners to the restored table
    reattachTableEventListeners();

    // Update climate summary
    if (updateClimateSummary) {
      updateClimateSummary();
    }

    setRestoringState(false);
  }

  /**
   * Save current state for undo/redo
   */
  function saveState() {
    // Don't save state if we're currently restoring
    if (isRestoringState) return;

    const state = {
      outputHTML: output.innerHTML,
      layerData: new Map(dataRefs.layerData),
      climateData: new Map(dataRefs.climateData),
      lastRows: dataRefs.lastRows ? JSON.parse(JSON.stringify(dataRefs.lastRows)) : null,
      lastHeaders: dataRefs.lastHeaders ? [...dataRefs.lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };

    undoStack.push(state);

    // Limit stack size
    if (undoStack.length > maxUndoSteps) {
      undoStack.shift();
    }

    // Clear redo stack when new action is performed
    redoStack.length = 0;

    updateUndoRedoButtons();
  }

  /**
   * Perform undo operation
   */
  function performUndo() {
    if (undoStack.length === 0) return;

    // Save current state to redo stack
    const currentState = {
      outputHTML: output.innerHTML,
      layerData: new Map(dataRefs.layerData),
      climateData: new Map(dataRefs.climateData),
      lastRows: dataRefs.lastRows ? JSON.parse(JSON.stringify(dataRefs.lastRows)) : null,
      lastHeaders: dataRefs.lastHeaders ? [...dataRefs.lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
    redoStack.push(currentState);

    // Restore previous state
    const previousState = undoStack.pop();
    restoreState(previousState);

    updateUndoRedoButtons();
  }

  /**
   * Perform redo operation
   */
  function performRedo() {
    if (redoStack.length === 0) return;

    // Save current state to undo stack
    const currentState = {
      outputHTML: output.innerHTML,
      layerData: new Map(dataRefs.layerData),
      climateData: new Map(dataRefs.climateData),
      lastRows: dataRefs.lastRows ? JSON.parse(JSON.stringify(dataRefs.lastRows)) : null,
      lastHeaders: dataRefs.lastHeaders ? [...dataRefs.lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
    undoStack.push(currentState);

    // Restore next state
    const nextState = redoStack.pop();
    restoreState(nextState);

    updateUndoRedoButtons();
  }

  return {
    saveState,
    performUndo,
    performRedo,
    updateUndoRedoButtons,
    restoreState
  };
}
