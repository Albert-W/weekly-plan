/**
 * Habits sheet logic (Google Sheets edition).
 *
 * Ported from src/taskpane/js/habits.js. Completion is recorded when the
 * habit's native checkbox (column B) is checked — see Triggers.gs onEdit
 * — instead of on cell selection. After recording, the checkbox is reset
 * so it behaves like a "mark done" button. Streak math is unchanged:
 *   weightedScore = base * 1.1 ^ streak
 */

/**
 * 1-based index of the last habit row (column A), or DATA_START_ROW - 1.
 * @returns {number}
 */
function getLastHabitRow_() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return CONFIG.HABITS.DATA_START_ROW - 1;
  return Math.max(getLastRowInColumn_(sheet, 1), CONFIG.HABITS.DATA_START_ROW - 1);
}

/**
 * Initialize the Habits sheet (find today's column, highlight it).
 */
function initializeHabitsSheet() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return;

  let dayIndex = findHabitsDayIndex(sheet);
  if (dayIndex < 0) {
    refreshHabitsDatesCore_(sheet);
    dayIndex = findHabitsDayIndex(sheet);
  }
  highlightCurrentDateHeader(sheet);
}

/**
 * Find today's column index (0-13) within the Habits header row, or -1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function findHabitsDayIndex(sheet) {
  const todayDay = new Date().getDate();
  const values = sheet.getRange(CONFIG.HABITS.HEADER_RANGE).getValues()[0];
  for (let i = 0; i < values.length; i++) {
    if (parseInt(values[i], 10) === todayDay) return i;
  }
  return -1;
}

/**
 * Highlight today's date column header (D3:Q3).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function highlightCurrentDateHeader(sheet) {
  const H = CONFIG.HABITS;
  sheet.getRange(H.HEADER_RANGE).setBackground(CONFIG.COLORS.CLEAR);

  const dayIndex = findHabitsDayIndex(sheet);
  if (dayIndex >= 0) {
    const colIndex = columnLetterToIndex(H.COLUMNS.DAY_START) + 1 + dayIndex; // 1-based
    sheet.getRange(H.HEADER_ROW, colIndex).setBackground(CONFIG.COLORS.TODAY_HIGHLIGHT);
  }
}

/**
 * Refresh the 14-day window: B3 = "yyyy mm", D3:Q3 = day numbers from
 * the current Monday, and clear the completion data area.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Date} the Monday the window starts from
 */
function refreshHabitsDatesCore_(sheet) {
  const H = CONFIG.HABITS;
  const startDate = getMonday(new Date());

  const yearMonth =
    startDate.getFullYear() + ' ' + String(startDate.getMonth() + 1).padStart(2, '0');
  sheet.getRange(H.YEAR_MONTH_CELL).setValue(yearMonth);

  const days = [];
  for (let i = 0; i < H.DAYS_COUNT; i++) {
    const d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    days.push(d.getDate());
  }
  sheet.getRange(H.HEADER_RANGE).setValues([days]);

  const lastRow = getLastHabitRow_();
  if (lastRow >= H.DATA_START_ROW) {
    const dayStartCol = columnLetterToIndex(H.COLUMNS.DAY_START) + 1;
    const dayColCount = H.DAYS_COUNT;
    sheet
      .getRange(H.DATA_START_ROW, dayStartCol, lastRow - H.DATA_START_ROW + 1, dayColCount)
      .clearContent();
  }
  return startDate;
}

/**
 * Public refresh-dates entry point (menu / sidebar).
 * @returns {string} status message
 */
function refreshHabitsDates() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) {
    toast_('Habits sheet not found.', 'Weekly Plan');
    return 'Habits sheet not found.';
  }
  const startDate = refreshHabitsDatesCore_(sheet);
  highlightCurrentDateHeader(sheet);
  const msg = 'Dates refreshed — window starts ' + startDate.toDateString();
  toast_(msg, 'Weekly Plan');
  return msg;
}

/**
 * Record a habit completion for a row (triggered by checking its box).
 * Applies a streak bonus and resets the checkbox afterwards.
 * @param {number} row 1-based row of the habit
 */
function recordHabitDone(row) {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return;
  const H = CONFIG.HABITS;

  const checkboxCol = columnLetterToIndex(H.COLUMNS.DONE_CHECKBOX) + 1;
  const resetCheckbox = function () {
    sheet.getRange(row, checkboxCol).setValue(false);
  };

  const dayIndex = findHabitsDayIndex(sheet);
  if (dayIndex < 0) {
    toast_('Today not found in the window. Click "Dates" to refresh.', 'Weekly Plan');
    resetCheckbox();
    return;
  }

  const habitName = sheet.getRange(H.COLUMNS.HABIT_NAME + row).getValue();
  if (!habitName) {
    resetCheckbox();
    return;
  }

  let weightedScore = 0;
  let streak = 0;
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('recordHabitDone: lock failed: ' + e.message);
    resetCheckbox();
    return;
  }

  try {
    const baseScore = parseFloat(sheet.getRange(H.COLUMNS.BASE_SCORE + row).getValue()) || 1;
    const dayStartCol = columnLetterToIndex(H.COLUMNS.DAY_START) + 1;
    const dayValues = sheet.getRange(row, dayStartCol, 1, H.DAYS_COUNT).getValues()[0];
    const totalCell = sheet.getRange(H.COLUMNS.TOTAL_COUNT + row);

    for (let d = dayIndex - 1; d >= 0; d--) {
      const v = dayValues[d];
      if (v && v !== 0) streak++;
      else break;
    }

    weightedScore = baseScore * Math.pow(H.STREAK_MULTIPLIER, streak);
    const currentCount = parseInt(dayValues[dayIndex], 10) || 0;
    const currentTotal = parseInt(totalCell.getValue(), 10) || 0;

    sheet.getRange(row, dayStartCol + dayIndex).setValue(currentCount + 1);
    totalCell.setValue(currentTotal + 1);
    sheet.getRange(H.COLUMNS.HABIT_NAME + row).setBackground(CONFIG.COLORS.POSITIVE);
    resetCheckbox();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  updateSummary(weightedScore, 0);

  const streakMsg = streak > 0 ? ' — ' + (streak + 1) + '-day streak!' : '';
  toast_('"' + habitName + '" +' + weightedScore.toFixed(2) + ' pts' + streakMsg, 'Weekly Plan');
}

/**
 * Sort habit rows by base score (column C) descending.
 * @returns {string} status message
 */
function sortHabits() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) {
    toast_('Habits sheet not found.', 'Weekly Plan');
    return 'Habits sheet not found.';
  }
  const H = CONFIG.HABITS;
  const lastRow = getLastHabitRow_();
  if (lastRow < H.DATA_START_ROW) return 'No habits to sort.';

  const baseScoreCol = columnLetterToIndex(H.COLUMNS.BASE_SCORE) + 1; // C -> 3
  const lastCol = columnLetterToIndex(H.COLUMNS.TOTAL_COUNT) + 1; // R
  sheet
    .getRange(H.DATA_START_ROW, 1, lastRow - H.DATA_START_ROW + 1, lastCol)
    .sort({ column: baseScoreCol, ascending: false });

  toast_('Habits sorted by score.', 'Weekly Plan');
  return 'Habits sorted by score.';
}
