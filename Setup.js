/**
 * Auto-scaffold for the Weekly Plan Google Sheets edition.
 *
 * This has no equivalent in the Office add-in (which assumed a
 * pre-built workbook). `setUpSheets()` builds the whole structure so
 * the user never has to recreate the intricate grid by hand:
 *   - Weekly: time-block grid, day headers, score dropdowns + color rules
 *   - Habits: 14-day window + native checkboxes
 *   - Tasks / Summary / Archive: headers
 *
 * Idempotent: re-running repairs structure (headers, time column,
 * validation, conditional formatting, checkboxes) without wiping the
 * user's task/habit/score data.
 */

/**
 * Build or repair every sheet the add-on needs.
 * @returns {string} human-readable summary
 */
function setUpSheets() {
  // Tasks and Habits must be built before Weekly so the combined
  // task+habit dropdown list (hidden _Dropdown sheet) is complete
  // when applyTaskValidation_ wires up the cell dropdowns.
  setUpTasksSheet_();
  setUpHabitsSheet_();
  setUpWeeklySheet_();
  setUpSummarySheet_();
  setUpDiarySheet_();
  setUpArchiveSheet_();
  toast_('Sheets are set up and ready.', 'Weekly Plan');
  return 'Weekly, Habits, Tasks, Summary, Diary, and Archive are ready.';
}

// ----------------------------------------------------------------------
// Weekly
// ----------------------------------------------------------------------

function setUpWeeklySheet_() {
  const sheet = getOrCreateSheet_(CONFIG.WEEKLY_SHEET);
  const W = CONFIG.WEEKLY;

  // Top control bar (clickable buttons handled by handleSelection).
  setUpControlButtons_(sheet);

  // Day-name labels (left, in task columns). Day numbers go in the
  // score columns and are written by setNewWeekDates (right-aligned).
  const dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
    sheet
      .getRange(W.HEADER_ROW, getTaskColForDay(d))
      .setValue(dayNames[d])
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    sheet
      .getRange(W.HEADER_ROW, getScoreColForDay(d))
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
  }

  // Time column B5:B{LAST_TIME_ROW} as decimal hours (8, 8.5, 9, …).
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;
  const times = [];
  for (let i = 0; i < slotCount; i++) {
    times.push([W.FIRST_HOUR + i * 0.5]);
  }
  sheet
    .getRange(W.DATA_START_ROW, W.TIME_COLUMN, slotCount, 1)
    // Force a plain decimal format — the column may carry a leftover
    // time format that would render 8 / 8.5 as 00:00 / 12:00.
    .setNumberFormat('0.#')
    .setValues(times)
    .setHorizontalAlignment('center');

  // Daily totals row label ("Scores" under the time column).
  sheet
    .getRange(W.SCORE_ROW, W.TIME_COLUMN)
    .setValue('Scores')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Score-cell dropdowns + color rules across all 7 day score columns.
  applyScoreValidationAndColors_(sheet);

  // Task-cell dropdowns sourced from the Tasks sheet (pick, don't type).
  applyTaskValidation_(sheet);

  // Populate the current week's dates (B4 year-month + day numbers).
  setNewWeekDates(sheet);

  // Grid borders to match the calendar look.
  setUpWeeklyBorders_(sheet);

  // Freeze the header + time column for usability.
  sheet.setFrozenRows(W.HEADER_ROW);
  sheet.setFrozenColumns(W.TIME_COLUMN);
}

/**
 * Refresh the hidden _Dropdown sheet with a combined list of task names
 * and habit names, so the Weekly task-cell dropdowns offer both.
 * Sorted by weight descending (task weight = col B, habit weight = col C),
 * then alphabetically for ties.  Idempotent: safe to call from setup,
 * task CRUD, and habit refresh.
 */
function refreshDropdownList_() {
  const hiddenSheet = getOrCreateSheet_(CONFIG.HIDDEN_SHEET);
  hiddenSheet.hideSheet();

  var entries = []; // { name:string, weight:number }

  // --- Tasks ---
  // Read name (col A) + weight (col B) in one batch.
  var tasksSheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (tasksSheet) {
    var lastRow = getLastTaskRow_();
    if (lastRow >= CONFIG.TASKS.DATA_START_ROW) {
      var taskRows = lastRow - CONFIG.TASKS.DATA_START_ROW + 1;
      var taskData = tasksSheet
        .getRange(CONFIG.TASKS.DATA_START_ROW, 1, taskRows, 2)
        .getValues();
      for (var i = 0; i < taskData.length; i++) {
        var name = taskData[i][0];
        if (name) {
          entries.push({
            name: String(name).trim(),
            weight: parseFloat(taskData[i][1]) || 0,
          });
        }
      }
    }
  }

  // --- Habits ---
  // Read name (col A) + weight (col C) in one batch.  Column B is the
  // "done" checkbox — we skip it but have to read through it because the
  // columns are contiguous.
  var habitsSheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (habitsSheet) {
    var lastRow = getLastHabitRow_();
    var H = CONFIG.HABITS;
    if (lastRow >= H.DATA_START_ROW) {
      var habitRows = lastRow - H.DATA_START_ROW + 1;
      var nameColIdx = columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1; // A -> 1
      // Read from name column through base-score column (col A..C).
      var habitData = habitsSheet
        .getRange(H.DATA_START_ROW, nameColIdx, habitRows, 3)
        .getValues();
      for (var i = 0; i < habitData.length; i++) {
        var name = habitData[i][0];
        if (name) {
          entries.push({
            name: String(name).trim(),
            weight: parseFloat(habitData[i][2]) || 0, // col C = index 2 in a 3-col read
          });
        }
      }
    }
  }

  // Sort: weight descending, then alphabetically for ties.
  entries.sort(function (a, b) {
    if (b.weight !== a.weight) return b.weight - a.weight;
    return a.name.localeCompare(b.name);
  });

  // Write just the names to the hidden sheet.
  hiddenSheet.clearContents();
  if (entries.length > 0) {
    var names = entries.map(function (e) { return [e.name]; });
    hiddenSheet.getRange(1, 1, names.length, 1).setValues(names);
  }
}

/**
 * Add a dropdown to every task cell, sourced from the hidden _Dropdown
 * sheet that combines task + habit names so both can be picked for a
 * time block. Invalid values are allowed so Random Pick fills and the
 * auto-created "others" task aren't flagged.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Weekly sheet
 */
function applyTaskValidation_(sheet) {
  const W = CONFIG.WEEKLY;

  // Ensure the combined list is fresh before we wire up the dropdown.
  refreshDropdownList_();

  const hiddenSheet = getOrCreateSheet_(CONFIG.HIDDEN_SHEET);
  // Include all possible rows so new entries auto-appear in the dropdown.
  const namesRange = hiddenSheet.getRange(1, 1, hiddenSheet.getMaxRows(), 1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(namesRange, true)
    .setAllowInvalid(true)
    .build();

  const firstDataRow = W.DATA_START_ROW;
  const rowCount = W.LAST_TIME_ROW - firstDataRow + 1;
  for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
    sheet
      .getRange(firstDataRow, getTaskColForDay(d), rowCount, 1)
      .setDataValidation(rule);
  }
}

/**
 * Draw the light-blue clickable control buttons in CONTROL_ROW.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setUpControlButtons_(sheet) {
  const W = CONFIG.WEEKLY;
  const buttons = W.CONTROL_BUTTONS;
  for (let i = 0; i < buttons.length; i++) {
    const b = buttons[i];
    const range = sheet.getRange(W.CONTROL_ROW, b.col, 1, 2);
    range
      .merge()
      .setValue(b.label)
      .setBackground(CONFIG.COLORS.BUTTON_FILL)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false);
  }
}

/**
 * Draw calendar-style borders: a box around the time column and each
 * day's task/score pair, plus a divider under the mid-day hour.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setUpWeeklyBorders_(sheet) {
  const W = CONFIG.WEEKLY;
  const firstRow = W.HEADER_ROW;
  const rowCount = W.SCORE_ROW - W.HEADER_ROW + 1;

  // Time column box.
  sheet
    .getRange(firstRow, W.TIME_COLUMN, rowCount, 1)
    .setBorder(true, true, true, true, false, false);

  // One box per day pair (task+score columns).
  for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
    sheet
      .getRange(firstRow, getTaskColForDay(d), rowCount, 2)
      .setBorder(true, true, true, true, false, false);
  }

  // Header underline + Scores-row overline across the whole grid.
  const lastCol = getScoreColForDay(W.DAYS_IN_WEEK - 1);
  const gridWidth = lastCol - W.TIME_COLUMN + 1;
  sheet.getRange(W.HEADER_ROW, W.TIME_COLUMN, 1, gridWidth).setBorder(
    null, null, true, null, null, null
  );
  sheet.getRange(W.SCORE_ROW, W.TIME_COLUMN, 1, gridWidth).setBorder(
    true, null, null, null, null, null
  );

  // Mid-day divider under MID_DIVIDER_HOUR (e.g. 17.5).
  const dividerOffset = Math.round((W.MID_DIVIDER_HOUR - W.FIRST_HOUR) / 0.5);
  const dividerRow = W.DATA_START_ROW + dividerOffset;
  if (dividerRow > W.DATA_START_ROW && dividerRow < W.SCORE_ROW) {
    sheet.getRange(dividerRow, W.TIME_COLUMN, 1, gridWidth).setBorder(
      null, null, true, null, null, null
    );
  }
}

/**
 * Add a 0/0.2/.../1 dropdown to every score cell and (re)build the
 * conditional-format color rules. Score coloring is therefore handled
 * by Sheets natively (survives sorts/recalc) instead of per-edit
 * setBackground as in the Office build.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function applyScoreValidationAndColors_(sheet) {
  const W = CONFIG.WEEKLY;
  const firstDataRow = W.DATA_START_ROW;
  const lastDataRow = W.SCORE_ROW - 1;

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(W.SCORE_OPTIONS.map(String), true)
    // Allow invalid so the numeric value written back by
    // processWeeklyScoreChange (which makes conditional formatting work)
    // isn't flagged against the string-list dropdown.
    .setAllowInvalid(true)
    .build();

  const scoreRanges = [];
  for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
    const col = getScoreColForDay(d);
    sheet
      .getRange(firstDataRow, col, lastDataRow - firstDataRow + 1, 1)
      .setDataValidation(validation);
    scoreRanges.push(sheet.getRange(firstDataRow, col, W.SCORE_ROW - firstDataRow + 1, 1));
  }

  // Rebuild conditional-format rules for the score columns only.
  // Keep any rules the user added on unrelated ranges.
  const existing = sheet.getConditionalFormatRules();
  const scoreA1 = scoreRanges.map((r) => r.getA1Notation());
  const kept = existing.filter((rule) => {
    const ranges = rule.getRanges().map((r) => r.getA1Notation());
    return !ranges.some((a1) => scoreA1.indexOf(a1) !== -1);
  });

  const positive = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground(CONFIG.COLORS.POSITIVE)
    .setRanges(scoreRanges)
    .build();
  const neutral = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground(CONFIG.COLORS.NEUTRAL)
    .setRanges(scoreRanges)
    .build();
  const negative = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground(CONFIG.COLORS.NEGATIVE)
    .setRanges(scoreRanges)
    .build();

  sheet.setConditionalFormatRules(kept.concat([positive, neutral, negative]));
}

// ----------------------------------------------------------------------
// Habits
// ----------------------------------------------------------------------

function setUpHabitsSheet_() {
  const sheet = getOrCreateSheet_(CONFIG.HABITS_SHEET);
  const H = CONFIG.HABITS;

  // Header row labels. Note: B3 (the checkbox column header) is set by
  // refreshHabitsDatesCore_ to the "yyyy mm" window label, so we don't
  // write a "Done?" label there.
  sheet.getRange(H.COLUMNS.HABIT_NAME + H.HEADER_ROW).setValue('Habit').setFontWeight('bold');
  sheet.getRange(H.COLUMNS.WEIGHT + H.HEADER_ROW).setValue('Weight').setFontWeight('bold');
  sheet.getRange(H.COLUMNS.TOTAL_COUNT + H.HEADER_ROW).setValue('Total').setFontWeight('bold');

  // Seed a few sample habits on first run only (don't clobber data).
  const lastNameRow = getLastRowInColumn_(sheet, columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1);
  if (lastNameRow < H.DATA_START_ROW) {
    const samples = [
      ['Exercise', 1],
      ['Read', 1],
      ['Meditate', 1],
    ];
    for (let i = 0; i < samples.length; i++) {
      const row = H.DATA_START_ROW + i;
      sheet.getRange(H.COLUMNS.HABIT_NAME + row).setValue(samples[i][0]);
      sheet.getRange(H.COLUMNS.WEIGHT + row).setValue(samples[i][1]);
    }
  }

  // Add the diary habit ("写日记") if missing. Runs even when sample seeding
  // was skipped — existing users already have habits and still need this one.
  ensureDiaryHabit_();

  // Native checkboxes in the "Done?" column for a generous row range so
  // habits typed later already have a checkbox. Empty rows are ignored
  // by recordHabitDone (it guards on habit name).
  const checkboxLastRow = Math.max(H.DATA_START_ROW + 49, lastNameRow);
  sheet
    .getRange(
      H.DATA_START_ROW,
      columnLetterToIndex(H.COLUMNS.DONE_CHECKBOX) + 1,
      checkboxLastRow - H.DATA_START_ROW + 1,
      1
    )
    .insertCheckboxes();

  // Populate the 14-day window header (B3 + D3:Q3) on first setup only.
  // Re-running must NOT clear existing completion data, so we check
  // whether the day headers already have values before refreshing.
  const firstDayHeader = sheet.getRange(H.HEADER_RANGE).getValues()[0];
  const hasDateHeaders = firstDayHeader.some(function (v) { return v !== '' && v !== null; });
  if (!hasDateHeaders) {
    refreshHabitsDatesCore_(sheet);
  }
  highlightCurrentDateHeader(sheet);

  sheet.setFrozenRows(H.HEADER_ROW);
}

/**
 * Add the "写日记" habit (weight 1) to the Habits sheet if not present.
 * Idempotent — safe on every set-up / re-run.
 */
function ensureDiaryHabit_() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return;
  const D = CONFIG.DIARY;
  if (findHabitRowByName_(D.HABIT_NAME) >= 0) return;
  const lastRow = getLastRowInColumn_(sheet, 1);
  const row = Math.max(lastRow + 1, CONFIG.HABITS.DATA_START_ROW);
  sheet.getRange(CONFIG.HABITS.COLUMNS.HABIT_NAME + row).setValue(D.HABIT_NAME);
  sheet.getRange(CONFIG.HABITS.COLUMNS.WEIGHT + row).setValue(D.HABIT_WEIGHT);
}

// ----------------------------------------------------------------------
// Tasks
// ----------------------------------------------------------------------

function setUpTasksSheet_() {
  const sheet = getOrCreateSheet_(CONFIG.TASKS_SHEET);
  const headerRow = CONFIG.TASKS.DATA_START_ROW - 1; // row 3

  sheet.getRange('A' + headerRow).setValue('Task').setFontWeight('bold');
  sheet.getRange('B' + headerRow).setValue('Weight').setFontWeight('bold');
  sheet.getRange('C' + headerRow).setValue('Created').setFontWeight('bold');
  sheet.getRange('D' + headerRow).setValue('Last Done').setFontWeight('bold');
  sheet.getRange('F' + headerRow).setValue('Count').setFontWeight('bold');
  sheet.getRange('G' + headerRow).setValue('Total Score').setFontWeight('bold');

  // Seed sample tasks on first run only.
  const lastTaskRow = getLastRowInColumn_(sheet, 1);
  if (lastTaskRow < CONFIG.TASKS.DATA_START_ROW) {
    const samples = [
      ['Deep Work', 1.5],
      ['Email', 0.5],
      ['Break', 1],
    ];
    for (let i = 0; i < samples.length; i++) {
      const row = CONFIG.TASKS.DATA_START_ROW + i;
      sheet.getRange('A' + row).setValue(samples[i][0]);
      sheet.getRange('B' + row).setValue(samples[i][1]);
      sheet.getRange('C' + row).setValue(formatDateTime(new Date()));
    }
  }

  sheet.setFrozenRows(headerRow);
}

// ----------------------------------------------------------------------
// Summary
// ----------------------------------------------------------------------

function setUpSummarySheet_() {
  const sheet = getOrCreateSheet_(CONFIG.SUMMARY_SHEET);
  const S = CONFIG.SUMMARY;
  sheet.getRange(S.DATE_COLUMN + '1').setValue('Date').setFontWeight('bold');
  sheet.getRange(S.POSITIVE_SCORE_COLUMN + '1').setValue('Positive').setFontWeight('bold');
  sheet.getRange(S.NEGATIVE_SCORE_COLUMN + '1').setValue('Negative').setFontWeight('bold');
  sheet.getRange(S.TOTAL_SCORE_COLUMN + '1').setValue('Total').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ----------------------------------------------------------------------
// Archive
// ----------------------------------------------------------------------

function setUpArchiveSheet_() {
  const sheet = getOrCreateSheet_(CONFIG.ARCHIVE_SHEET);
  if (getLastRowInColumn_(sheet, 1) >= 1) return; // header already present

  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const header = ['Week', 'Time'];
  for (let d = 0; d < days.length; d++) {
    header.push(days[d] + '_Task');
    header.push(days[d] + '_Score');
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
