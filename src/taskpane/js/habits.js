/**
 * Habits sheet logic for the Combined Tracker Add-in
 *
 * This file contains all functions specific to the Habits sheet,
 * including initialization, recording completions, and streak calculations.
 */

/**
 * Initialize Habits sheet data
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeHabitsSheet(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);
  const usedRange = sheet.getUsedRange();
  usedRange.load('rowCount');
  await context.sync();

  state.habits.lastRow = Math.max(usedRange.rowCount, CONFIG.HABITS.DATA_START_ROW);
  state.habits.currentDayIndex = await findHabitsDayIndex(context, sheet);

  // If today's date is not found in the table, refresh the dates
  if (state.habits.currentDayIndex < 0) {
    console.log('Today\'s date not found in habits table, refreshing dates...');
    await refreshHabitsDatesWithContext(context, sheet);
    state.habits.currentDayIndex = await findHabitsDayIndex(context, sheet);
  }

  // Highlight the current date column header
  await highlightCurrentDateHeader(context, sheet);

  console.log('Habits initialized:', state.habits);
}

/**
 * Highlight the current date column header in the Habits sheet
 * Clears previous highlighting and applies new highlight to today's date
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Habits sheet
 */
async function highlightCurrentDateHeader(context, sheet) {
  // Clear all date header highlighting first (D3:Q3)
  const headerRange = sheet.getRange('D3:Q3');
  headerRange.format.fill.color = CONFIG.COLORS.CLEAR;

  // If we have a valid current day index, highlight that column header
  if (state.habits.currentDayIndex >= 0) {
    const dayColIndex = columnLetterToIndex(CONFIG.HABITS.COLUMNS.DAY_START) + state.habits.currentDayIndex;
    const dayColumn = indexToColumnLetter(dayColIndex);
    const currentDateCell = sheet.getRange(`${dayColumn}${CONFIG.HABITS.HEADER_ROW}`);
    currentDateCell.format.fill.color = CONFIG.COLORS.TODAY_HIGHLIGHT;
  }

  await context.sync();
}

/**
 * Refresh dates for Habits sheet using an existing context.
 * Sets B3 = "yyyy mm", D3:Q3 = day numbers for the 14-day window
 * starting at the Monday of `today`, and clears any prior data.
 *
 * Single source of truth for the date-refresh logic. Callers that
 * already own an Excel.RequestContext call this directly; the
 * public refreshHabitsDates() below wraps it in Excel.run.
 *
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Habits sheet
 * @returns {Promise<Date>} The Monday the window starts from
 */
async function refreshHabitsDatesWithContext(context, sheet) {
  const today = new Date();
  const startDate = getMonday(today);

  // Set year/month in B3
  const yearMonth = `${startDate.getFullYear()} ${String(startDate.getMonth() + 1).padStart(2, '0')}`;
  sheet.getRange('B3').values = [[yearMonth]];

  // Set day numbers in header row
  const days = [];
  for (let i = 0; i < CONFIG.HABITS.DAYS_COUNT; i++) {
    const d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    days.push(d.getDate());
  }
  sheet.getRange('D3:Q3').values = [days];

  // Clear data area
  if (state.habits.lastRow >= CONFIG.HABITS.DATA_START_ROW) {
    sheet.getRange(`D${CONFIG.HABITS.DATA_START_ROW}:Q${state.habits.lastRow}`).clear(Excel.ClearApplyTo.contents);
  }

  await context.sync();
  console.log('Dates refreshed! Starting ' + startDate.toDateString());
  return startDate;
}

/**
 * Find current day index in the Habits header row
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Habits sheet
 * @returns {number} Day index (0-13) or -1 if not found
 */
async function findHabitsDayIndex(context, sheet) {
  const today = new Date();
  const todayDay = today.getDate();

  const headerRange = sheet.getRange('D3:Q3');
  headerRange.load('values');
  await context.sync();

  const values = headerRange.values[0];
  for (let i = 0; i < values.length; i++) {
    if (parseInt(values[i]) === todayDay) {
      return i;
    }
  }
  return -1;
}

/**
 * Handle selection on Habits sheet
 * @param {Excel.RequestContext} context - Excel context
 * @param {string} address - Cell address
 * @param {string} column - Column letter
 * @param {number} colIndex - Column index (1-based)
 * @param {number} row - Row number
 */
async function handleHabitsSelection(context, address, column, colIndex, row) {
  // Click on habit name (Column A) in data area → record completion
  if (column === 'A' && row >= CONFIG.HABITS.DATA_START_ROW && row <= state.habits.lastRow) {
    await recordHabitDone(context, row);
    return;
  }

  // Click on A2 → show help
  if (column === 'A' && row === 2) {
    toggleSection('habits-help');
    return;
  }
}

/**
 * Record habit completion for a row
 * Calculates streak bonus and updates counts.
 *
 * Performance: batches reads/writes into 2 context.sync() calls
 * (one before computing, one after writing) regardless of streak
 * length, instead of N+5 round-trips.
 *
 * @param {Excel.RequestContext} context - Excel context
 * @param {number} row - Row number of the habit
 */
async function recordHabitDone(context, row) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);

  if (state.habits.currentDayIndex < 0) {
    showStatus('Current date not found. Click Refresh Dates.', 'error');
    return;
  }

  // ------------------------------------------------------------------
  // 1. Queue all reads we need, then sync ONCE.
  // ------------------------------------------------------------------
  const habitCell = sheet.getRange(`A${row}`);
  const baseScoreCell = sheet.getRange(`C${row}`);
  const dayStart = CONFIG.HABITS.COLUMNS.DAY_START;
  const dayEnd = CONFIG.HABITS.COLUMNS.DAY_END;
  const daysRange = sheet.getRange(`${dayStart}${row}:${dayEnd}${row}`);
  const totalCell = sheet.getRange(`${CONFIG.HABITS.COLUMNS.TOTAL_COUNT}${row}`);

  habitCell.load('values');
  baseScoreCell.load('values');
  daysRange.load('values');
  totalCell.load('values');
  await context.sync();

  const habitName = habitCell.values[0][0];
  if (!habitName) return;

  // ------------------------------------------------------------------
  // 2. Compute streak and weighted score from in-memory values.
  // ------------------------------------------------------------------
  const dayValues = daysRange.values[0]; // single row -> array of 14 cells
  const todayIndex = state.habits.currentDayIndex;

  let streak = 0;
  for (let d = todayIndex - 1; d >= 0; d--) {
    const v = dayValues[d];
    if (v && v !== 0) {
      streak++;
    } else {
      break;
    }
  }

  const baseScore = parseFloat(baseScoreCell.values[0][0]) || 1;
  const weightedScore = baseScore * Math.pow(CONFIG.HABITS.STREAK_MULTIPLIER, streak);

  const currentCount = parseInt(dayValues[todayIndex]) || 0;
  const currentTotal = parseInt(totalCell.values[0][0]) || 0;

  // ------------------------------------------------------------------
  // 3. Queue all writes, then sync ONCE.
  // ------------------------------------------------------------------
  const dayColIndex = columnLetterToIndex(dayStart) + todayIndex;
  const dayColumn = indexToColumnLetter(dayColIndex);

  sheet.getRange(`${dayColumn}${row}`).values = [[currentCount + 1]];
  totalCell.values = [[currentTotal + 1]];
  sheet.getRange(`B${row}`).format.fill.color = CONFIG.COLORS.POSITIVE;

  await context.sync();

  // Update summary sheet (this owns its own sync calls)
  await updateSummary(context, weightedScore, 0);

  const streakMsg = streak > 0 ? ` 🔥 ${streak + 1}-day streak!` : '';
  showStatus(`✅ "${habitName}" +${weightedScore.toFixed(2)} pts${streakMsg}`, 'success');
}

/**
 * Sort habits by base score (descending)
 */
async function sortHabits() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);
      const range = sheet.getRange(`A${CONFIG.HABITS.DATA_START_ROW}:R${state.habits.lastRow}`);

      range.sort.apply([{ key: 2, ascending: false }]);
      await context.sync();

      showStatus('Habits sorted by score!', 'success');
    });
  } catch (error) {
    showStatus('Error sorting: ' + error.message, 'error');
  }
}

/**
 * Refresh dates for Habits sheet (public, user-facing entry point).
 * Thin wrapper that owns Excel.run and updates the current-day
 * index + status. All sheet work lives in
 * refreshHabitsDatesWithContext.
 */
async function refreshHabitsDates() {
  try {
    let startDate;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);
      startDate = await refreshHabitsDatesWithContext(context, sheet);
      state.habits.currentDayIndex = await findHabitsDayIndex(context, sheet);
    });
    showStatus('Dates refreshed! Starting ' + startDate.toDateString(), 'success');
  } catch (error) {
    showStatus('Error: ' + error.message, 'error');
  }
}

// Export for use in other modules
window.initializeHabitsSheet = initializeHabitsSheet;
window.findHabitsDayIndex = findHabitsDayIndex;
window.handleHabitsSelection = handleHabitsSelection;
window.recordHabitDone = recordHabitDone;
window.sortHabits = sortHabits;
window.refreshHabitsDates = refreshHabitsDates;
