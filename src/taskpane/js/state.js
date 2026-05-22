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
    currentDayIndex: 0,     // Current day index (0-6 for Mon-Sun)
    lastMonday: null,
    lastTaskRow: 4,         // Last row in Tasks sheet
    lastSummaryRow: 1,      // Last row in Summary sheet
    lastInitDate: null,     // Track last initialization date (YYYY-MM-DD)
    // Grid extent. Initialized from CONFIG defaults and overwritten
    // by initializeWeeklySheet once we know what the actual sheet
    // looks like (task #9 — sheet-driven layout).
    lastTimeRow: CONFIG.WEEKLY.LAST_TIME_ROW,
    scoreRow: CONFIG.WEEKLY.SCORE_ROW
  },

  // Event handlers
  selectionHandler: null,
  changeHandler: null
};

// Export for use in other modules
window.state = state;
