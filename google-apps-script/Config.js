/**
 * Configuration constants for the Weekly Plan Google Sheets edition.
 *
 * Ported from src/taskpane/js/config.js. In Apps Script every .gs file
 * shares one global scope, so this top-level `const CONFIG` is visible
 * to every other module — no `window.*` export needed.
 */

const CONFIG = {
  // ==================== SHEET NAMES ====================
  HABITS_SHEET: 'Habits',
  WEEKLY_SHEET: 'Weekly',
  TASKS_SHEET: 'Tasks',
  SUMMARY_SHEET: 'Summary',
  ARCHIVE_SHEET: 'Archive', // GAS-only: in-spreadsheet history of finished weeks

  // ==================== HABITS CONFIG ====================
  HABITS: {
    DATA_START_ROW: 4,
    HEADER_ROW: 3,
    COLUMNS: {
      HABIT_NAME: 'A',
      DONE_CHECKBOX: 'B', // GAS-only: native checkbox column ("mark done")
      BASE_SCORE: 'C',
      DAY_START: 'D',
      DAY_END: 'Q',
      TOTAL_COUNT: 'R',
    },
    YEAR_MONTH_CELL: 'B3',
    HEADER_RANGE: 'D3:Q3',
    DAYS_COUNT: 14,
    STREAK_MULTIPLIER: 1.1,
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
    TIME_COLUMN: 2, // Column B for timestamps
    DATE_CELL: 'B4', // Cell containing "yyyy mm" format
    FIRST_DAY_HEADER_CELL: 'D4', // First day-number header (Monday)
    HEADER_RANGE: 'D4:P4', // All 7 day-number headers (D, F, H, J, L, N, P)
    HEADER_ROW_RANGE: 'A4:P4', // Whole header row, for fill clears
    LAST_TIME_ROW: 36, // Last row with time data
    SCORE_ROW: 38, // Score totals row
    // Task columns (odd: 3,5,7,9,11,13,15) = C,E,G,I,K,M,O
    TASK_COLUMNS: [3, 5, 7, 9, 11, 13, 15],
    // Score columns (even: 4,6,8,10,12,14,16) = D,F,H,J,L,N,P
    SCORE_COLUMNS: [4, 6, 8, 10, 12, 14, 16],
    // Score options for the data-validation dropdown
    SCORE_OPTIONS: [0, 0.2, 0.4, 0.6, 0.8, 1],
    // Days in week
    DAYS_IN_WEEK: 7,
    // Scaffold: first hour and last hour for the time-block grid.
    // Rows DATA_START_ROW..LAST_TIME_ROW hold one 30-min slot each,
    // shown as decimal hours (8, 8.5, 9, …) to match the original.
    FIRST_HOUR: 8, // 08:00
    // Mid-day visual divider drawn under this decimal hour (e.g. 17.5).
    MID_DIVIDER_HOUR: 17.5,
    // In-sheet control buttons in CONTROL_ROW. Each is merged across
    // `col`..`col+1`. `action` maps to a handler in handleSelection (Triggers).
    CONTROL_BUTTONS: [
      { action: 'help', label: 'Help', col: 3 }, // C2:D2
      { action: 'add', label: 'Add Task', col: 5 }, // E2:F2
      { action: 'delete', label: 'Delete Task', col: 7 }, // G2:H2
      { action: 'random', label: 'Random Fill', col: 9 }, // I2:J2
      { action: 'thanks', label: 'Thanks', col: 11 }, // K2:L2
    ],
  },

  // ==================== SUMMARY CONFIG ====================
  SUMMARY: {
    DATE_COLUMN: 'A',
    POSITIVE_SCORE_COLUMN: 'D',
    NEGATIVE_SCORE_COLUMN: 'E',
    TOTAL_SCORE_COLUMN: 'F',
  },

  // ==================== COLORS ====================
  COLORS: {
    TODAY_HIGHLIGHT: '#FFFF00', // Yellow
    POSITIVE: '#70AD47', // Green
    NEGATIVE: '#ED7D31', // Orange-Red
    NEUTRAL: '#FFC000', // Yellow/amber
    CURRENT_TIME: '#FFFF00', // Yellow for current hour
    CLEAR: '#FFFFFF',
    BUTTON_FILL: '#DCE6F1', // Light blue control-bar buttons
  },

  // ==================== DRIVE / ARCHIVE ====================
  DRIVE_ARCHIVE_FOLDER: 'Weekly Plan Archives',
};
