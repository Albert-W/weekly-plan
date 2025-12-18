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

  console.log('Habits initialized:', state.habits);
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
 * Calculates streak bonus and updates counts
 * @param {Excel.RequestContext} context - Excel context
 * @param {number} row - Row number of the habit
 */
async function recordHabitDone(context, row) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);

  if (state.habits.currentDayIndex < 0) {
    showStatus('Current date not found. Click Refresh Dates.', 'error');
    return;
  }

  // Get habit name
  const habitCell = sheet.getRange(`A${row}`);
  habitCell.load('values');
  await context.sync();

  const habitName = habitCell.values[0][0];
  if (!habitName) return;

  // Calculate day column
  const dayColIndex = columnLetterToIndex(CONFIG.HABITS.COLUMNS.DAY_START) + state.habits.currentDayIndex;
  const dayColumn = indexToColumnLetter(dayColIndex);

  // Calculate streak (consecutive days before today)
  let streak = 0;
  for (let d = state.habits.currentDayIndex - 1; d >= 0; d--) {
    const prevCol = indexToColumnLetter(columnLetterToIndex(CONFIG.HABITS.COLUMNS.DAY_START) + d);
    const prevCell = sheet.getRange(`${prevCol}${row}`);
    prevCell.load('values');
    await context.sync();

    if (prevCell.values[0][0] && prevCell.values[0][0] !== 0) {
      streak++;
    } else {
      break;
    }
  }

  // Get base score
  const scoreCell = sheet.getRange(`C${row}`);
  scoreCell.load('values');
  await context.sync();
  const baseScore = parseFloat(scoreCell.values[0][0]) || 1;

  // Calculate weighted score with streak bonus
  const weightedScore = baseScore * Math.pow(CONFIG.HABITS.STREAK_MULTIPLIER, streak);

  // Increment day count
  const dayCell = sheet.getRange(`${dayColumn}${row}`);
  dayCell.load('values');
  await context.sync();
  const currentCount = parseInt(dayCell.values[0][0]) || 0;
  dayCell.values = [[currentCount + 1]];

  // Update total count
  const totalCell = sheet.getRange(`R${row}`);
  totalCell.load('values');
  await context.sync();
  const total = parseInt(totalCell.values[0][0]) || 0;
  totalCell.values = [[total + 1]];

  // Highlight the label cell
  sheet.getRange(`B${row}`).format.fill.color = CONFIG.COLORS.POSITIVE;

  await context.sync();

  // Update summary sheet
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
 * Refresh dates for Habits sheet
 * Sets up a new 14-day window starting from Monday
 */
async function refreshHabitsDates() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);

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

      // Update current day index
      state.habits.currentDayIndex = await findHabitsDayIndex(context, sheet);
      showStatus('Dates refreshed! Starting ' + startDate.toDateString(), 'success');
    });
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
