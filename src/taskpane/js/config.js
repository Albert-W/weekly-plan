/**
 * Configuration constants for the Combined Tracker Add-in
 *
 * This file contains all configuration values for both
 * Habits and Weekly/Timetable sheets.
 */

const CONFIG = {
  // Sheet names
  HABITS_SHEET: 'Habits',
  WEEKLY_SHEET: 'Weekly',
  TIMETABLE_SHEET: 'Timetable',  // Alternative name for Weekly
  TASKS_SHEET: 'Tasks',
  SUMMARY_SHEET: 'Summary',
  GOALS_SHEET: 'Goals',
  CHARTER_SHEET: 'Charter',

  // ==================== HABITS CONFIG ====================
  HABITS: {
    DATA_START_ROW: 4,
    HEADER_ROW: 3,
    COLUMNS: {
      HABIT_NAME: 'A',
      HABIT_LABEL: 'B',
      BASE_SCORE: 'C',
      DAY_START: 'D',
      DAY_END: 'Q',
      TOTAL_COUNT: 'R'
    },
    // Derived address strings (kept here so other files never have
    // to hardcode 'B3' or 'D3:Q3').
    YEAR_MONTH_CELL: 'B3',
    HEADER_RANGE: 'D3:Q3',
    DAYS_COUNT: 14,
    STREAK_MULTIPLIER: 1.1
  },

  // ==================== TASKS CONFIG ====================
  TASKS: {
    // Name used for the auto-created "catch-all" row when a scored
    // task isn't found in the Tasks list.
    FALLBACK_NAME: 'others',
    // First row of actual task data (rows 1-3 are headers).
    DATA_START_ROW: 4,
  },

  // ==================== WEEKLY/TIMETABLE CONFIG ====================
  WEEKLY: {
    DATA_START_ROW: 5,
    CONTROL_ROW: 2,
    HEADER_ROW: 4,
    TIME_COLUMN: 2,     // Column B for timestamps
    DATE_CELL: 'B4',    // Cell containing "yyyy mm" format
    FIRST_DAY_HEADER_CELL: 'D4',  // First day-number header (Monday)
    HEADER_RANGE: 'D4:P4',        // All 7 day-number headers (D, F, H, J, L, N, P)
    HEADER_ROW_RANGE: 'A4:P4',    // Whole header row, for fill clears
    LAST_TIME_ROW: 36,  // Last row with time data
    SCORE_ROW: 38,      // Score totals row
    // Control buttons in row 2
    BUTTONS: {
      HELP: 3,      // Column C
      ADD: 5,       // Column E
      RANDOM: 9,    // Column I
    },
    // Task columns (odd: 3,5,7,9,11,13,15) = C,E,G,I,K,M,O
    TASK_COLUMNS: [3, 5, 7, 9, 11, 13, 15],
    // Score columns (even: 4,6,8,10,12,14,16) = D,F,H,J,L,N,P
    SCORE_COLUMNS: [4, 6, 8, 10, 12, 14, 16],
    // Score options for dropdown
    SCORE_OPTIONS: [0, 0.2, 0.4, 0.6, 0.8, 1],
    // Days in week
    DAYS_IN_WEEK: 7
  },

  // ==================== SUMMARY CONFIG ====================
  SUMMARY: {
    DATE_COLUMN: 'A',
    POSITIVE_SCORE_COLUMN: 'D',
    NEGATIVE_SCORE_COLUMN: 'E',
    TOTAL_SCORE_COLUMN: 'F'
  },

  // ==================== COLORS ====================
  COLORS: {
    TODAY_HIGHLIGHT: '#FFFF00',    // Yellow (65535 in VBA)
    POSITIVE: '#70AD47',           // Green (Accent6)
    NEGATIVE: '#ED7D31',           // Orange-Red (Accent2)
    NEUTRAL: '#FFC000',            // Yellow (Accent5)
    CURRENT_TIME: '#FFFF00',       // Yellow for current hour
    CLEAR: '#FFFFFF'
  }
};

// Export for use in other modules (works in browser)
window.CONFIG = CONFIG;
