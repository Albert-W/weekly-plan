/**
 * Weekly/Timetable sheet logic for the Combined Tracker Add-in
 *
 * This file contains all functions specific to the Weekly sheet,
 * including initialization, random pick, score processing, and time highlighting.
 */

/**
 * Initialize Weekly/Timetable sheet data
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeWeeklySheet(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

  // Find last row with time data (Column B)
  const timeColumn = sheet.getRange('B:B').getUsedRange();
  timeColumn.load('rowCount');
  await context.sync();

  // Using fixed values from CONFIG instead of dynamic calculation
  // CONFIG.WEEKLY.LAST_TIME_ROW = 36, CONFIG.WEEKLY.SCORE_ROW = 38

  // Calculate current day index (0=Mon, 6=Sun)
  const today = new Date();
  const dayOfWeek = today.getDay();
  state.weekly.currentDayIndex = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

  // Calculate last Monday
  state.weekly.lastMonday = getMonday(today);

  console.log('Weekly initialized:', state.weekly);
}

/**
 * Initialize Weekly sheet on workbook open
 * Equivalent to VBA Workbook_Open logic
 *
 * If today is in a new week (7+ days since last Monday in sheet):
 *   1. Archive the current week data (export as CSV)
 *   2. Clear the table for new week
 *   3. Set new week dates
 *
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeWeeklyOnOpen(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

  // Get date from B4 (format: "yyyy mm")
  const dateCell = sheet.getRange(CONFIG.WEEKLY.DATE_CELL);
  dateCell.load('values');

  // Get first day number from D4
  const firstDayCell = sheet.getRange(CONFIG.WEEKLY.FIRST_DAY_HEADER_CELL);
  firstDayCell.load('values');

  // Find last row with time data first (needed for archive)
  const timeColumn = sheet.getRange('B:B').getUsedRange();
  timeColumn.load('rowCount');

  await context.sync();

  // Using fixed values from CONFIG: LAST_TIME_ROW = 36, SCORE_ROW = 38

  const dateStr = String(dateCell.values[0][0] || '');
  const firstDay = parseInt(firstDayCell.values[0][0]) || 0;

  console.log('Weekly date cell:', dateStr, 'First day:', firstDay);

  // Parse the date from sheet
  let sheetLastMonday = null;
  if (dateStr) {
    const parts = dateStr.split(' ');
    if (parts.length >= 2) {
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1; // JavaScript months are 0-based
      sheetLastMonday = new Date(year, month, firstDay);
    }
  }

  // Calculate days since lastMonday
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (sheetLastMonday) {
    const diffDays = daysBetween(sheetLastMonday, today);
    console.log('Days since last Monday:', diffDays);

    // If 7+ days have passed, it's a NEW WEEK!
    if (diffDays >= 7) {
      console.log('🗓️ New week detected! Archiving previous week...');
      showStatus('🗓️ New week detected! Archiving...', 'info');

      // Step 1: Archive the current week data BEFORE clearing
      await archiveWeekAutomatically();

      // Step 2: Clear the table for new week
      await clearForNewWeek(context);

      // Step 3: Set new week dates
      await setNewWeekDates(context, sheet);

      showStatus('✅ Week archived & new week started!', 'success');
    }
  } else {
    // No valid date, set new dates (first time use)
    console.log('No valid date found, setting up first week...');
    await setNewWeekDates(context, sheet);
  }

  // Using fixed CONFIG values for LAST_TIME_ROW and SCORE_ROW

  // Calculate current day index (0=Mon, 6=Sun)
  const dayOfWeek = today.getDay();
  state.weekly.currentDayIndex = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

  // Calculate last Monday
  state.weekly.lastMonday = getMonday(today);

  // Highlight current day column
  await highlightCurrentDay(context, sheet);

  // Highlight current time row
  await highlightCurrentTimeRow(context, sheet);

  // Track the initialization date
  state.weekly.lastInitDate = formatDateYYYYMMDD(today);

  console.log('Weekly fully initialized:', state.weekly);
}

async function setNewWeekDates(context, sheet) {
  const today = new Date();
  const newMonday = getMonday(today);

  // Set B4 = "yyyy mm"
  const yearMonth = `${newMonday.getFullYear()} ${String(newMonday.getMonth() + 1).padStart(2, '0')}`;
  sheet.getRange(CONFIG.WEEKLY.DATE_CELL).values = [[yearMonth]];

  // Set day numbers in D4, F4, H4, J4, L4, N4, P4
  for (let i = 0; i < CONFIG.WEEKLY.DAYS_IN_WEEK; i++) {
    const dayDate = new Date(newMonday);
    dayDate.setDate(newMonday.getDate() + i);
    const colLetter = getScoreColLetterForDay(i);
    sheet.getRange(`${colLetter}4`).values = [[dayDate.getDate()]];
  }

  state.weekly.lastMonday = newMonday;

  await context.sync();
  console.log('New week dates set, starting:', newMonday);
}

/**
 * Clear content for new week.
 * Equivalent to VBA clearForNewWeek().
 *
 * Behavior preserved: only clear the task+score pair for rows that
 * actually had a score recorded. Rows with a task selected but no
 * score yet are left untouched (matches the legacy VBA).
 *
 * Performance: previously ran 7 days x ~33 rows = ~231 individual
 * context.sync() calls reading one score cell at a time. Now it
 * loads the whole C5:P{SCORE_ROW-1} data area in a single read,
 * scans the in-memory 2D array, queues all clears at once, and
 * commits with a single final sync. Total: 2 syncs.
 *
 * @param {Excel.RequestContext} context - Excel context
 */
async function clearForNewWeek(context) {
  try {
    const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
    const dataStart = CONFIG.WEEKLY.DATA_START_ROW;
    const dataEnd = CONFIG.WEEKLY.SCORE_ROW - 1;

    // ----------------------------------------------------------------
    // 1. Queue unconditional clears + the bulk read, then sync ONCE.
    //    All of these ride on the same round-trip as the values read.
    // ----------------------------------------------------------------
    sheet.getRange(`C5:Z${CONFIG.WEEKLY.SCORE_ROW}`).format.fill.clear();
    sheet
      .getRange(`C${CONFIG.WEEKLY.SCORE_ROW}:P${CONFIG.WEEKLY.SCORE_ROW}`)
      .clear(Excel.ClearApplyTo.contents);

    const dataRange = sheet.getRange(`C${dataStart}:P${dataEnd}`);
    dataRange.load('values');
    await context.sync();

    // ----------------------------------------------------------------
    // 2. Scan values in memory. dataRange.values is a 2D array indexed
    //    [rowOffset][colOffset]. Column C is offset 0, so day d's
    //    task column = d*2, score column = d*2 + 1.
    // ----------------------------------------------------------------
    const values = dataRange.values;
    for (let day = 0; day < CONFIG.WEEKLY.DAYS_IN_WEEK; day++) {
      const taskColOffset = day * 2;
      const scoreColOffset = day * 2 + 1;
      const taskColLetter = getTaskColLetterForDay(day);
      const scoreColLetter = getScoreColLetterForDay(day);

      for (let i = 0; i < values.length; i++) {
        const scoreVal = values[i][scoreColOffset];
        if (scoreVal !== '' && scoreVal !== null) {
          const row = dataStart + i;
          sheet.getRange(`${taskColLetter}${row}`).clear(Excel.ClearApplyTo.contents);
          sheet.getRange(`${scoreColLetter}${row}`).clear(Excel.ClearApplyTo.contents);
        }
      }
    }

    // ----------------------------------------------------------------
    // 3. Commit every queued clear in one round-trip.
    // ----------------------------------------------------------------
    await context.sync();
    console.log('Cleared content for new week');
    showStatus('Cleared for new week!', 'success');
  } catch (error) {
    console.error('Clear for new week error:', error);
  }
}

/**
 * Highlight current day column (header row)
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Weekly sheet
 */
async function highlightCurrentDay(context, sheet) {
  // Clear previous highlighting in header row
  sheet.getRange(CONFIG.WEEKLY.HEADER_ROW_RANGE).format.fill.clear();

  // Highlight current day's task and score header columns
  const taskColLetter = getTaskColLetterForDay(state.weekly.currentDayIndex);
  const scoreColLetter = getScoreColLetterForDay(state.weekly.currentDayIndex);

  sheet.getRange(`${taskColLetter}4`).format.fill.color = CONFIG.COLORS.TODAY_HIGHLIGHT;
  sheet.getRange(`${scoreColLetter}4`).format.fill.color = CONFIG.COLORS.TODAY_HIGHLIGHT;

  await context.sync();
  console.log('Highlighted current day column:', state.weekly.currentDayIndex);
}

/**
 * Highlight current time row
 * Equivalent to VBA hourTask()
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Weekly sheet
 */
async function highlightCurrentTimeRow(context, sheet) {
  try {
    // Clear previous time highlighting
    sheet.getRange('B5:B' + CONFIG.WEEKLY.LAST_TIME_ROW).format.fill.clear();

    // Get current time as decimal (e.g., 15.75 for 15:45)
    const now = new Date();
    const currentHour = now.getHours();
    const currentMinutes = now.getMinutes();
    const currentTimeDecimal = currentHour + currentMinutes / 60;

    console.log('=== HIGHLIGHT TIME DEBUG ===');
    console.log('Current time:', currentHour + ':' + currentMinutes, '= decimal:', currentTimeDecimal);
    console.log('Looking in rows', CONFIG.WEEKLY.DATA_START_ROW, 'to', CONFIG.WEEKLY.LAST_TIME_ROW);

    // Get time column values
    const timeRange = sheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.LAST_TIME_ROW}`);
    timeRange.load('values');
    await context.sync();

    console.log('Time range values count:', timeRange.values.length);

    // Find the best matching time row
    let bestRowIndex = -1;
    let bestTimeValue = -1;

    for (let i = 0; i < timeRange.values.length; i++) {
      const cellTime = timeRange.values[i][0];
      if (cellTime === '' || cellTime === null || cellTime === undefined) continue;

      let timeValue = null;

      // Determine the type of time value
      if (typeof cellTime === 'number') {
        // Excel time: could be fraction of day (0.0-1.0) or hours (8, 9, 15.5, etc.)
        if (cellTime >= 0 && cellTime <= 1) {
          // Fraction of day (e.g., 0.645833 = 15:30)
          timeValue = cellTime * 24;
        } else if (cellTime >= 0 && cellTime <= 24) {
          // Already in hours (e.g., 15.5 = 15:30)
          timeValue = cellTime;
        } else {
          // Some other number, skip
          continue;
        }
      } else if (typeof cellTime === 'string') {
        // Try to parse "HH:MM" or "H:MM" format
        const match = cellTime.match(/^(\d{1,2}):(\d{2})$/);
        if (match) {
          timeValue = parseInt(match[1]) + parseInt(match[2]) / 60;
        } else {
          continue;
        }
      } else {
        continue;
      }

      const rowNum = CONFIG.WEEKLY.DATA_START_ROW + i;
    //   console.log('Row', rowNum, ': cellTime=', cellTime, '(type:', typeof cellTime, ') -> timeValue=', timeValue.toFixed(2));

      // We want to find the time slot that contains the current time
      // Match if timeValue is within 0.5 hour before current time, up to current time + small buffer
      // This finds the current/most recent time slot
      if (timeValue <= currentTimeDecimal + 0.1 && timeValue > bestTimeValue) {
        bestTimeValue = timeValue;
        bestRowIndex = i;
      }
    }

    console.log('Best match: rowIndex=', bestRowIndex, 'timeValue=', bestTimeValue);

    // Highlight the best matching row
    if (bestRowIndex >= 0) {
      const row = CONFIG.WEEKLY.DATA_START_ROW + bestRowIndex;
      console.log('>>> Highlighting row:', row);

      // Highlight time cell
      sheet.getRange(`B${row}`).format.fill.color = CONFIG.COLORS.CURRENT_TIME;

      // Check if current day's task/score cells are empty
      const taskColLetter = getTaskColLetterForDay(state.weekly.currentDayIndex);
      const scoreColLetter = getScoreColLetterForDay(state.weekly.currentDayIndex);

      const scoreCell = sheet.getRange(`${scoreColLetter}${row}`);
      scoreCell.load('values');
      await context.sync();

      // If no score entered, highlight task and score cells too
      if (scoreCell.values[0][0] === '' || scoreCell.values[0][0] === null) {
        sheet.getRange(`${taskColLetter}${row}`).format.fill.color = CONFIG.COLORS.CURRENT_TIME;
        sheet.getRange(`${scoreColLetter}${row}`).format.fill.color = CONFIG.COLORS.CURRENT_TIME;
      }

      await context.sync();
      console.log('Highlighted row', row, 'for time', bestTimeValue.toFixed(2));
    } else {
      console.log('No matching time row found!');
    }

    console.log('=== END HIGHLIGHT TIME DEBUG ===');
  } catch (error) {
    console.error('Highlight time row error:', error);
  }
}

/**
 * Refresh current time highlighting.
 * @param {Object} [opts]
 * @param {boolean} [opts.silent=false] - When true, suppress the status banner.
 *   Used by the auto-tick so the user isn't pinged every minute.
 */
async function refreshTimeHighlight({ silent = false } = {}) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.WEEKLY_SHEET);
      await context.sync();

      if (!sheet.isNullObject) {
        await highlightCurrentTimeRow(context, sheet);
      }
    });
    if (!silent) showStatus('Time highlight refreshed!', 'success');
  } catch (error) {
    if (!silent) showStatus('Error refreshing time: ' + error.message, 'error');
    else console.error('Time highlight tick error:', error);
  }
}

// ----------------------------------------------------------------------
// Background ticker: keep the highlighted time row in sync with the
// real clock so the user sees the highlight move when each new minute
// (and therefore each new time slot) begins. Only does Excel work when
// the user is currently on the Weekly sheet.
// ----------------------------------------------------------------------

let _timeHighlightTimerId = null;
let _timeHighlightAlignTimeoutId = null;

/**
 * Run one tick of the auto-highlight. Exported for testing.
 */
async function tickTimeHighlight() {
  if (state.currentSheet !== CONFIG.WEEKLY_SHEET) return;
  await refreshTimeHighlight({ silent: true });
}

/**
 * Start the background ticker. Idempotent — safe to call multiple
 * times. Aligns to the next minute boundary so the highlight flips
 * shortly after the clock minute changes, then runs every 60s.
 */
function startTimeHighlightTicker() {
  if (_timeHighlightTimerId || _timeHighlightAlignTimeoutId) return;
  const now = new Date();
  const msUntilNextMinute = (60 - now.getSeconds()) * 1000 - now.getMilliseconds();
  _timeHighlightAlignTimeoutId = setTimeout(() => {
    _timeHighlightAlignTimeoutId = null;
    tickTimeHighlight();
    _timeHighlightTimerId = setInterval(tickTimeHighlight, 60_000);
  }, msUntilNextMinute);
}

/**
 * Stop the background ticker (used for cleanup / tests).
 */
function stopTimeHighlightTicker() {
  if (_timeHighlightAlignTimeoutId) {
    clearTimeout(_timeHighlightAlignTimeoutId);
    _timeHighlightAlignTimeoutId = null;
  }
  if (_timeHighlightTimerId) {
    clearInterval(_timeHighlightTimerId);
    _timeHighlightTimerId = null;
  }
}

/**
 * Handle selection on Weekly sheet
 * @param {Excel.RequestContext} context - Excel context
 * @param {string} address - Cell address
 * @param {string} column - Column letter
 * @param {number} colIndex - Column index (1-based)
 * @param {number} row - Row number
 */
async function handleWeeklySelection(context, address, column, colIndex, row) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

  // Control row buttons (Row 2)
  if (row === CONFIG.WEEKLY.CONTROL_ROW) {
    switch (colIndex) {
      case CONFIG.WEEKLY.BUTTONS.HELP:
        toggleSection('weekly-help');
        break;
      case CONFIG.WEEKLY.BUTTONS.ADD:
        toggleSection('add-task-section');
        break;
      case CONFIG.WEEKLY.BUTTONS.DELETE:
        toggleSection('delete-task-section');
        break;
      case CONFIG.WEEKLY.BUTTONS.RANDOM:
        await randomPick(context);
        break;
      case CONFIG.WEEKLY.BUTTONS.THANK:
        toggleSection('thank-section');
        break;
    }
    return;
  }

  // Task column selection (odd columns 3-15, rows 5+)
  if (CONFIG.WEEKLY.TASK_COLUMNS.includes(colIndex) &&
      row >= CONFIG.WEEKLY.DATA_START_ROW && row <= CONFIG.WEEKLY.LAST_TIME_ROW) {
    showStatus('Select a task from the dropdown, or use Add Task to create new.', 'info');
    return;
  }

  // Score column selection (even columns 4-16, rows 5+)
  if (CONFIG.WEEKLY.SCORE_COLUMNS.includes(colIndex) &&
      row >= CONFIG.WEEKLY.DATA_START_ROW && row < CONFIG.WEEKLY.SCORE_ROW) {

    // Check if task is selected first
    const taskCell = sheet.getRange(address).getOffsetRange(0, -1);
    taskCell.load('values');
    await context.sync();

    if (!taskCell.values[0][0]) {
      showWarningPopup('Please select a task first!');
      return;
    }

    // Check if already has score
    const scoreCell = sheet.getRange(address);
    scoreCell.load('values');
    await context.sync();

    if (scoreCell.values[0][0] !== '' && scoreCell.values[0][0] !== null) {
      showStatus("Score can't be modified once set.", 'warning');
      return;
    }

    showStatus('Enter a score (0, 0.2, 0.4, 0.6, 0.8, or 1)', 'info');
  }
}

/**
 * Random Pick - fill empty task slots with random tasks
 * Equivalent to VBA RandomPick()
 * @param {Excel.RequestContext} context - Excel context
 */
async function randomPick(context) {
  try {
    const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
    const tasksSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.TASKS_SHEET);
    await context.sync();

    if (tasksSheet.isNullObject) {
      showStatus('Tasks sheet not found!', 'error');
      return;
    }

    // Get all tasks (starting from row 4)
    const tasksRange = tasksSheet.getRange(`A${CONFIG.TASKS.DATA_START_ROW}:A${state.weekly.lastTaskRow}`);
    tasksRange.load('values');
    await context.sync();

    const tasks = tasksRange.values.filter(t => t[0] && t[0] !== '');
    if (tasks.length === 0) {
      showStatus('No tasks available for random pick', 'error');
      return;
    }

    // Current day task column
    const taskColLetter = getTaskColLetterForDay(state.weekly.currentDayIndex);

    // Get time column and task column
    const timeRange = weeklySheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.LAST_TIME_ROW}`);
    const taskRange = weeklySheet.getRange(`${taskColLetter}${CONFIG.WEEKLY.DATA_START_ROW}:${taskColLetter}${CONFIG.WEEKLY.LAST_TIME_ROW}`);

    timeRange.load('values');
    taskRange.load('values');
    await context.sync();

    let filledCount = 0;

    for (let i = 0; i < timeRange.values.length; i++) {
      const hasTime = timeRange.values[i][0] !== '' && timeRange.values[i][0] !== null;
      const hasTask = taskRange.values[i][0] !== '' && taskRange.values[i][0] !== null;

      if (hasTime && !hasTask) {
        const randomIndex = Math.floor(Math.random() * tasks.length);
        const randomTask = tasks[randomIndex][0];
        const row = CONFIG.WEEKLY.DATA_START_ROW + i;
        weeklySheet.getRange(`${taskColLetter}${row}`).values = [[randomTask]];
        filledCount++;
      }
    }

    await context.sync();

    if (filledCount > 0) {
      showStatus(`🎲 Filled ${filledCount} slots with random tasks!`, 'success');
    } else {
      showStatus('No empty task slots with timestamps found', 'info');
    }
  } catch (error) {
    console.error('RandomPick error:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

/**
 * Process score change on Weekly sheet
 * Equivalent to VBA Worksheet_Change.
 *
 * Performance: batches Tasks-sheet lookups into a single
 * context.sync() instead of one round-trip per lookup, so total
 * syncs are constant (2) regardless of how many tasks exist.
 *
 * @param {Excel.RequestContext} context - Excel context
 * @param {number} row - Row number
 * @param {number} col - Column number (1-based)
 * @param {number} newScore - The new score value
 */
async function processWeeklyScoreChange(context, row, col, newScore) {
  const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
  const tasksSheet = context.workbook.worksheets.getItem(CONFIG.TASKS_SHEET);

  const taskColLetter = indexToColumnLetter(col - 2);
  const scoreColLetter = indexToColumnLetter(col - 1);

  // ------------------------------------------------------------------
  // 1. Queue every read we need, then sync ONCE.
  //    Loading whole columns (A, B, F, G) costs the same round-trip as
  //    loading any single cell, but lets us resolve the task in memory.
  // ------------------------------------------------------------------
  const taskCell = weeklySheet.getRange(`${taskColLetter}${row}`);
  const scoreCell = weeklySheet.getRange(`${scoreColLetter}${row}`);
  const dailyTotalCell = weeklySheet.getRange(
    `${scoreColLetter}${CONFIG.WEEKLY.SCORE_ROW}`
  );
  const tasksNames = tasksSheet.getRange(`A${CONFIG.TASKS.DATA_START_ROW}:A${state.weekly.lastTaskRow}`);
  const tasksWeights = tasksSheet.getRange(`B${CONFIG.TASKS.DATA_START_ROW}:B${state.weekly.lastTaskRow}`);
  const tasksCounts = tasksSheet.getRange(`F${CONFIG.TASKS.DATA_START_ROW}:F${state.weekly.lastTaskRow}`);
  const tasksScores = tasksSheet.getRange(`G${CONFIG.TASKS.DATA_START_ROW}:G${state.weekly.lastTaskRow}`);

  taskCell.load('values');
  dailyTotalCell.load('values');
  tasksNames.load('values');
  tasksWeights.load('values');
  tasksCounts.load('values');
  tasksScores.load('values');
  await context.sync();

  const taskName = taskCell.values[0][0];
  if (!taskName) return;

  // ------------------------------------------------------------------
  // 2. Resolve task in memory (no Excel calls).
  // ------------------------------------------------------------------
  let taskIndex = -1;     // 1-based row in Tasks sheet
  let othersIndex = -1;
  for (let i = 0; i < tasksNames.values.length; i++) {
    const name = tasksNames.values[i][0];
    if (name === taskName) {
      taskIndex = i + 4;
      break;
    }
    if (name === CONFIG.TASKS.FALLBACK_NAME) {
      othersIndex = i + 4;
    }
  }

  let isNewTask = false;
  let lookupIndex; // 0-based index into the loaded arrays
  if (taskIndex !== -1) {
    lookupIndex = taskIndex - 4;
  } else if (othersIndex !== -1) {
    taskIndex = othersIndex;
    lookupIndex = othersIndex - 4;
  } else {
    taskIndex = state.weekly.lastTaskRow + 1;
    lookupIndex = -1;
    isNewTask = true;
  }

  const taskWeight = isNewTask
    ? 1
    : parseFloat(tasksWeights.values[lookupIndex][0]) || 1;
  const currentCount = isNewTask
    ? 0
    : parseInt(tasksCounts.values[lookupIndex][0]) || 0;
  const currentTaskScore = isNewTask
    ? 0
    : parseFloat(tasksScores.values[lookupIndex][0]) || 0;

  const weightedScore = taskWeight * newScore;
  const currentDailyTotal = parseFloat(dailyTotalCell.values[0][0]) || 0;
  const newDailyTotal = currentDailyTotal + weightedScore;

  const color =
    weightedScore > 0 ? CONFIG.COLORS.POSITIVE :
    weightedScore < 0 ? CONFIG.COLORS.NEGATIVE :
    CONFIG.COLORS.NEUTRAL;

  const now = formatDateTime(new Date());

  // ------------------------------------------------------------------
  // 3. Queue every write, then sync ONCE.
  // ------------------------------------------------------------------
  scoreCell.format.fill.color = color;
  taskCell.format.fill.color = color;
  dailyTotalCell.values = [[newDailyTotal]];

  if (isNewTask) {
    // First-ever auto-creation of "others": populate every stat column
    // so the row is fully consistent.
    tasksSheet.getRange(`A${taskIndex}`).values = [[CONFIG.TASKS.FALLBACK_NAME]];
    tasksSheet.getRange(`B${taskIndex}`).values = [[1]];
    tasksSheet.getRange(`C${taskIndex}`).values = [[now]];
    tasksSheet.getRange(`D${taskIndex}`).values = [[now]];
    tasksSheet.getRange(`F${taskIndex}`).values = [[1]];
    tasksSheet.getRange(`G${taskIndex}`).values = [[weightedScore]];
    state.weekly.lastTaskRow = taskIndex;
  } else {
    tasksSheet.getRange(`D${taskIndex}`).values = [[now]];
    tasksSheet.getRange(`F${taskIndex}`).values = [[currentCount + 1]];
    tasksSheet.getRange(`G${taskIndex}`).values = [[currentTaskScore + weightedScore]];
  }

  await context.sync();

  // Update summary sheet (owns its own sync calls)
  await updateSummary(
    context,
    weightedScore > 0 ? weightedScore : 0,
    weightedScore < 0 ? weightedScore : 0
  );

  showStatus(
    `📝 "${taskName}" scored: ${weightedScore.toFixed(2)} (Daily: ${newDailyTotal.toFixed(2)})`,
    'success'
  );
}

async function startNewWeekFromUI() {
  await withStatus('Start new week', async () => {
    showStatus('🗓️ Starting new week...', 'info');
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await clearForNewWeek(context);
      await setNewWeekDates(context, sheet);
      await highlightCurrentDay(context, sheet);
      await highlightCurrentTimeRow(context, sheet);
    });
    showStatus('✅ New week started! Remember to save a copy for archive.', 'success');
  });
}


// Export for use in other modules
window.initializeWeeklySheet = initializeWeeklySheet;
window.initializeWeeklyOnOpen = initializeWeeklyOnOpen;
window.setNewWeekDates = setNewWeekDates;
window.clearForNewWeek = clearForNewWeek;
window.highlightCurrentDay = highlightCurrentDay;
window.highlightCurrentTimeRow = highlightCurrentTimeRow;
window.refreshTimeHighlight = refreshTimeHighlight;
window.handleWeeklySelection = handleWeeklySelection;
window.randomPick = randomPick;
window.processWeeklyScoreChange = processWeeklyScoreChange;
window.startNewWeekFromUI = startNewWeekFromUI;
window.tickTimeHighlight = tickTimeHighlight;
window.startTimeHighlightTicker = startTimeHighlightTicker;
window.stopTimeHighlightTicker = stopTimeHighlightTicker;

// ----------------------------------------------------------------------
// Sheet-handler registration (dispatched from events.js)
// ----------------------------------------------------------------------

/**
 * onActivate handler: when the Weekly sheet is activated, run the
 * new-day re-init if needed, otherwise just refresh the time-row
 * highlight.
 */
async function handleWeeklyActivate(context) {
  const today = formatDateYYYYMMDD(new Date());
  const lastInit = state.weekly.lastInitDate;
  console.log('Weekly sheet activated. Today:', today, 'Last init:', lastInit);

  if (lastInit !== today) {
    console.log('🌅 New day detected! Re-initializing Weekly sheet...');
    await initializeWeeklyOnOpen(context);
  } else {
    const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
    await highlightCurrentTimeRow(context, weeklySheet);
  }
}

/**
 * onChange handler: only act on score-column writes inside the data
 * area; everything else is ignored.
 */
async function handleWeeklyCellChange(context, address, colIndex, row) {
  if (
    CONFIG.WEEKLY.SCORE_COLUMNS.includes(colIndex) &&
    row >= CONFIG.WEEKLY.DATA_START_ROW &&
    row <= CONFIG.WEEKLY.LAST_TIME_ROW
  ) {
    const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
    const scoreCell = sheet.getRange(address);
    scoreCell.load('values');
    await context.sync();
    const scoreValue = parseFloat(scoreCell.values[0][0]);
    if (!isNaN(scoreValue)) {
      await processWeeklyScoreChange(context, row, colIndex, scoreValue);
    }
  }
}

if (typeof registerSheetHandlers === 'function') {
  registerSheetHandlers(CONFIG.WEEKLY_SHEET, {
    onSelection: handleWeeklySelection,
    onActivate: handleWeeklyActivate,
    onChange: handleWeeklyCellChange,
  });
}

window.handleWeeklyActivate = handleWeeklyActivate;
window.handleWeeklyCellChange = handleWeeklyCellChange;
