/**
 * Global state management for the Combined Tracker Add-in
 *
 * This file contains the shared state object that tracks
 * the current sheet, habits data, and weekly data.
 */

const state = {
  // Current sheet being viewed
  currentSheet: null,

  // Habits state
  habits: {
    currentDayIndex: -1,
    lastRow: 4
  },

  // Weekly state
  weekly: {
    dIndex: 0,          // Current day index (0-6 for Mon-Sun)
    lastMonday: null,
    taskl: 4,           // Last row in Tasks sheet
    summaryL: 1         // Last row in Summary sheet
  },

  // Event handlers
  selectionHandler: null
};

// Export for use in other modules
window.state = state;
