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
  setUpTasksSheet_(); // before Weekly so task dropdowns can reference it
  setUpWeeklySheet_();
  setUpHabitsSheet_();
  setUpSummarySheet_();
  setUpArchiveSheet_();
  toast_('Sheets are set up and ready.', 'Weekly Plan');
  return 'Weekly, Habits, Tasks, Summary, and Archive are ready.';
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
 * Add a dropdown to every task cell, sourced from the Tasks sheet's
 * name column, so tasks are picked rather than typed. Invalid values
 * are allowed so Random Pick fills and the auto-created "others" task
 * aren't flagged, and the list auto-updates as tasks are added/removed.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Weekly sheet
 */
function applyTaskValidation_(sheet) {
  const W = CONFIG.WEEKLY;
  const tasksSheet = getOrCreateSheet_(CONFIG.TASKS_SHEET);
  const start = CONFIG.TASKS.DATA_START_ROW;
  const namesRange = tasksSheet.getRange(start, 1, tasksSheet.getMaxRows() - start + 1, 1);

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
  sheet.getRange(H.COLUMNS.BASE_SCORE + H.HEADER_ROW).setValue('Base').setFontWeight('bold');
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
      sheet.getRange(H.COLUMNS.BASE_SCORE + row).setValue(samples[i][1]);
    }
  }

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

  // Populate the 14-day window header (B3 + D3:Q3) and current-day highlight.
  refreshHabitsDatesCore_(sheet);
  highlightCurrentDateHeader(sheet);

  sheet.setFrozenRows(H.HEADER_ROW);
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
