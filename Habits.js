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
 * 1-based row of a habit by name (column A), or -1 when not found.
 * Shared by Setup (ensureDiaryHabit_) and Diary (processPendingDiaryHabits_).
 * @param {string} name
 * @returns {number}
 */
function findHabitRowByName_(name) {
  if (!name) return -1;
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return -1;
  const H = CONFIG.HABITS;
  const last = getLastHabitRow_();
  if (last < H.DATA_START_ROW) return -1;
  const vals = sheet.getRange(H.DATA_START_ROW, 1, last - H.DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(name)) return H.DATA_START_ROW + i;
  }
  return -1;
}

/**
 * DocumentProperties key holding the YYYYMMDD date on which the habit
 * "done" fills were last reset. Used so the green completion color clears
 * exactly once per new day, not on every sidebar open.
 */
const HABITS_COLOR_RESET_PROP_ = 'habitsColorResetDate';

/**
 * Initialize the Habits sheet (find today's column, highlight it).
 * On the first run of a new day, also clears yesterday's green "done"
 * fills so every habit looks fresh and can be built again.
 */
function initializeHabitsSheet() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return;

  let dayIndex = findHabitsDayIndex(sheet);
  if (dayIndex < 0) {
    refreshHabitsDatesCore_(sheet);
    dayIndex = findHabitsDayIndex(sheet);
  }

  const props = PropertiesService.getDocumentProperties();
  const today = formatDateYYYYMMDD(new Date());
  if (props.getProperty(HABITS_COLOR_RESET_PROP_) !== today) {
    clearHabitDoneColors_(sheet);
    props.setProperty(HABITS_COLOR_RESET_PROP_, today);
  }

  highlightCurrentDateHeader(sheet);
}

/**
 * Clear the green "done" fill from the habit-name column (A) for every
 * habit row, so a new day starts visually fresh.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearHabitDoneColors_(sheet) {
  const H = CONFIG.HABITS;
  const lastRow = getLastHabitRow_();
  if (lastRow < H.DATA_START_ROW) return;
  const nameCol = columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1; // A -> 1
  sheet
    .getRange(H.DATA_START_ROW, nameCol, lastRow - H.DATA_START_ROW + 1, 1)
    .setBackground(CONFIG.COLORS.CLEAR);
}

/**
 * Find the window column index (0-13) for a given date, or -1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} date
 * @returns {number}
 */
function findHabitsDayIndexForDate_(sheet, date) {
  const day = date.getDate();
  const values = sheet.getRange(CONFIG.HABITS.HEADER_RANGE).getValues()[0];
  for (let i = 0; i < values.length; i++) {
    if (parseInt(values[i], 10) === day) return i;
  }
  return -1;
}

/**
 * Find today's column index (0-13) within the Habits header row, or -1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function findHabitsDayIndex(sheet) {
  return findHabitsDayIndexForDate_(sheet, new Date());
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
 * Record a habit completion for a row (triggered by checking its box, or
 * by the deferred diary-habit processing for a specific date).
 * Applies a streak bonus and resets the checkbox afterwards.
 * @param {number} row 1-based row of the habit
 * @param {Date} [date] date to mark (default: today)
 */
function recordHabitDone(row, date) {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return;
  const H = CONFIG.HABITS;

  const habitDate = date || new Date();
  const checkboxCol = columnLetterToIndex(H.COLUMNS.DONE_CHECKBOX) + 1;
  const resetCheckbox = function () {
    sheet.getRange(row, checkboxCol).setValue(false);
  };

  const dayIndex = findHabitsDayIndexForDate_(sheet, habitDate);
  if (dayIndex < 0) {
    toast_('Date not found in the window. Click "Dates" to refresh.', 'Weekly Plan');
    resetCheckbox();
    return;
  }

  const habitName = sheet.getRange(H.COLUMNS.HABIT_NAME + row).getValue();
  if (!habitName) {
    resetCheckbox();
    return;
  }

  // Mutable state set inside the lock, read afterwards for side effects.
  let weightedScore = 0;
  let streak = 0;
  let comboDays = 0;
  let isQuestHabit = false;

  // ------------------------------------------------------------------
  // Read-modify-write under the document lock. The quest combo is
  // read AND advanced atomically inside this lock (via
  // applyComboLocked_) so a concurrent habit/score handler can't read
  // a stale combo state and apply the wrong multiplier.
  // ------------------------------------------------------------------
  let finished = false;
  try {
    withLock_(function () {
      const weight = parseFloat(sheet.getRange(H.COLUMNS.WEIGHT + row).getValue()) || 1;
      const dayStartCol = columnLetterToIndex(H.COLUMNS.DAY_START) + 1;
      const dayValues = sheet.getRange(row, dayStartCol, 1, H.DAYS_COUNT).getValues()[0];
      const totalCell = sheet.getRange(H.COLUMNS.TOTAL_COUNT + row);

      for (let d = dayIndex - 1; d >= 0; d--) {
        const v = dayValues[d];
        if (v && v !== 0) streak++;
        else break;
      }

      weightedScore = weight * Math.pow(H.STREAK_MULTIPLIER, streak);
      const qHabitMult = questHabitMultiplier_(habitName); // Daily Quest bonus
      if (qHabitMult > 1) {
        // Atomically read AND advance the combo inside this lock
        // so the multiplier matches the state we persist.
        const combo = applyComboLocked_();
        comboDays = combo.days;
        weightedScore *= qHabitMult * combo.multiplier;
        isQuestHabit = true;
      } else {
        weightedScore *= qHabitMult;
      }
      const currentCount = parseInt(dayValues[dayIndex], 10) || 0;
      const currentTotal = parseInt(totalCell.getValue(), 10) || 0;

      sheet.getRange(row, dayStartCol + dayIndex).setValue(currentCount + 1);
      totalCell.setValue(currentTotal + 1);
      sheet.getRange(H.COLUMNS.HABIT_NAME + row).setBackground(CONFIG.COLORS.POSITIVE);
      resetCheckbox();
      finished = true;
    });
  } catch (e) {
    Logger.log('recordHabitDone: ' + (e && e.message ? e.message : e));
    resetCheckbox();
    return;
  }

  if (!finished) return;

  // Flush AFTER releasing the lock so we don't hold it during I/O.
  SpreadsheetApp.flush();

  updateSummary(weightedScore, 0);

  awardXp_(weightedScore);
  maybeAwardStreakBadge_(streak + 1);
  maybeAwardEarlyBird_(new Date());
  checkBossDefeat_();

  // Mark quest done (acquires its own lock via tryWithLock_).
  let comboNote = '';
  if (isQuestHabit) {
    markQuestDone_('habit', habitName);
    if (comboDays > 1) comboNote = ' 🔥' + comboDays + 'd combo';
  }

  const streakMsg = streak > 0 ? ' — ' + (streak + 1) + '-day streak!' : '';
  const questMsg = isQuestHabit ? ' ⭐ Quest bonus!' : '';
  toast_('"' + habitName + '" +' + weightedScore.toFixed(2) + ' pts' + streakMsg + questMsg + comboNote, 'Weekly Plan');
}

/**
 * Sort habit rows by weight (column C) descending.
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

  const weightCol = columnLetterToIndex(H.COLUMNS.WEIGHT) + 1; // C -> 3
  const lastCol = columnLetterToIndex(H.COLUMNS.TOTAL_COUNT) + 1; // R
  sheet
    .getRange(H.DATA_START_ROW, 1, lastRow - H.DATA_START_ROW + 1, lastCol)
    .sort({ column: weightCol, ascending: false });

  toast_('Habits sorted by weight.', 'Weekly Plan');
  return 'Habits sorted by weight.';
}
