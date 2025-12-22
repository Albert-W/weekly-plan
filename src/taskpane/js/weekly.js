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
  // CONFIG.WEEKLY.lastTimeLine = 36, CONFIG.WEEKLY.scoreLine = 38

  // Calculate current day index (0=Mon, 6=Sun)
  const today = new Date();
  const dayOfWeek = today.getDay();
  state.weekly.dIndex = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

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
  const firstDayCell = sheet.getRange('D4');
  firstDayCell.load('values');

  // Find last row with time data first (needed for archive)
  const timeColumn = sheet.getRange('B:B').getUsedRange();
  timeColumn.load('rowCount');

  await context.sync();

  // Using fixed values from CONFIG: lastTimeLine = 36, scoreLine = 38

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
      await archiveWeekAutomatically(sheetLastMonday, firstDay);

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

  // Using fixed CONFIG values for lastTimeLine and scoreLine

  // Calculate current day index (0=Mon, 6=Sun)
  const dayOfWeek = today.getDay();
  state.weekly.dIndex = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

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

/**
 * Automatically archive week data when new week is detected
 * This is called from initializeWeeklyOnOpen
 * @param {Date} weekMonday - The Monday of the week being archived
 * @param {number} firstDay - The first day number from the sheet
 */
async function archiveWeekAutomatically(weekMonday, firstDay) {
  try {
    console.log('Auto-archiving week starting:', weekMonday);

    let csvContent = '';
    let weekLabel = '';

    await Excel.run(async (context) => {
      const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

      // Get date info from B4
      const dateCell = weeklySheet.getRange('B4');
      dateCell.load('values');

      // Get day headers (D4:P4)
      const headerRange = weeklySheet.getRange('D4:P4');
      headerRange.load('values');

      // Get time column
      const timeRange = weeklySheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.lastTimeLine}`);
      timeRange.load('values');

      // Get all task and score data (C5:P[timel])
      const dataRange = weeklySheet.getRange(`C${CONFIG.WEEKLY.DATA_START_ROW}:P${CONFIG.WEEKLY.lastTimeLine}`);
      dataRange.load('values');

      await context.sync();

      // Build week label for filename
      const dateStr = String(dateCell.values[0][0] || '');
      const firstDayNum = headerRange.values[0][0];
      const lastDayNum = headerRange.values[0][headerRange.values[0].length - 1];
      weekLabel = `${dateStr.replace(' ', '-')}_${firstDayNum}-${lastDayNum}`;

      // Build CSV header
      const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
      let headers = ['Time'];
      for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
        headers.push(`${days[d]}_Task`);
        headers.push(`${days[d]}_Score`);
      }
      csvContent = headers.join(',') + '\n';

      // Build CSV rows
      for (let i = 0; i < timeRange.values.length; i++) {
        const time = timeRange.values[i][0];
        if (time === '' || time === null) continue;

        // Format time
        let timeStr;
        if (typeof time === 'number') {
          const hours = Math.floor(time * 24);
          const mins = Math.round((time * 24 - hours) * 60);
          timeStr = `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
        } else {
          timeStr = String(time);
        }

        let row = [escapeCSV(timeStr)];

        // Add task/score pairs for each day
        for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
          const taskCol = d * 2;      // 0, 2, 4, 6, 8, 10, 12
          const scoreCol = d * 2 + 1; // 1, 3, 5, 7, 9, 11, 13
          const task = dataRange.values[i][taskCol] || '';
          const score = dataRange.values[i][scoreCol];
          row.push(escapeCSV(String(task)));
          row.push(score !== null && score !== '' ? score : '');
        }

        csvContent += row.join(',') + '\n';
      }
    });

    // Download the CSV
    if (csvContent && csvContent.split('\n').length > 2) {
      const filename = `Weekly_${weekLabel}.csv`;
      downloadCSV(csvContent, filename);
      console.log('✅ Week archived to:', filename);
    } else {
      console.log('ℹ️ No data to archive for this week');
    }

  } catch (error) {
    console.error('Auto-archive error:', error);
    // Don't throw - allow the new week to start even if archive fails
  }
}

/**
 * Initialize Tasks sheet data
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeTasksSheet(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.TASKS_SHEET);
  const usedRange = sheet.getRange('A:A').getUsedRange();
  usedRange.load('rowCount');
  await context.sync();

  state.weekly.taskl = usedRange.rowCount;
  console.log('Tasks sheet initialized, taskl:', state.weekly.taskl);
}

/**
 * Initialize Summary sheet data
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeSummarySheet(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.SUMMARY_SHEET);
  const usedRange = sheet.getRange('A:A').getUsedRange();
  usedRange.load('rowCount');
  await context.sync();

  state.weekly.summaryL = usedRange.rowCount;
  console.log('Summary sheet initialized, summaryL:', state.weekly.summaryL);
}

/**
 * Set new week dates in the Weekly sheet
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Weekly sheet
 */
async function setNewWeekDates(context, sheet) {
  const today = new Date();
  const newMonday = getMonday(today);

  // Set B4 = "yyyy mm"
  const yearMonth = `${newMonday.getFullYear()} ${String(newMonday.getMonth() + 1).padStart(2, '0')}`;
  sheet.getRange('B4').values = [[yearMonth]];

  // Set day numbers in D4, F4, H4, J4, L4, N4, P4 (columns 4,6,8,10,12,14,16)
  for (let i = 0; i < CONFIG.WEEKLY.DAYS_IN_WEEK; i++) {
    const dayDate = new Date(newMonday);
    dayDate.setDate(newMonday.getDate() + i);
    const col = i * 2 + 4; // 4,6,8,10,12,14,16
    const colLetter = indexToColumnLetter(col - 1);
    sheet.getRange(`${colLetter}4`).values = [[dayDate.getDate()]];
  }

  state.weekly.lastMonday = newMonday;

  await context.sync();
  console.log('New week dates set, starting:', newMonday);
}

/**
 * Clear content for new week
 * Equivalent to VBA clearForNewWeek()
 * @param {Excel.RequestContext} context - Excel context
 */
async function clearForNewWeek(context) {
  try {
    const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

    // Clear background colors for data area
    const colorRange = sheet.getRange(`C5:Z${CONFIG.WEEKLY.scoreLine}`);
    colorRange.format.fill.clear();

    // Clear scores row
    const scoresRange = sheet.getRange(`C${CONFIG.WEEKLY.scoreLine}:P${CONFIG.WEEKLY.scoreLine}`);
    scoresRange.clear(Excel.ClearApplyTo.contents);

    // Clear each day's task and score cells
    for (let day = 0; day < CONFIG.WEEKLY.DAYS_IN_WEEK; day++) {
      const taskCol = day * 2 + 3;  // 3,5,7,9,11,13,15
      const scoreCol = day * 2 + 4; // 4,6,8,10,12,14,16
      const taskColLetter = indexToColumnLetter(taskCol - 1);
      const scoreColLetter = indexToColumnLetter(scoreCol - 1);

      for (let row = CONFIG.WEEKLY.DATA_START_ROW; row < CONFIG.WEEKLY.scoreLine; row++) {
        // Check if score cell has value
        const scoreCell = sheet.getRange(`${scoreColLetter}${row}`);
        scoreCell.load('values');
        await context.sync();

        if (scoreCell.values[0][0] !== '' && scoreCell.values[0][0] !== null) {
          // Clear task and score
          sheet.getRange(`${taskColLetter}${row}`).clear(Excel.ClearApplyTo.contents);
          sheet.getRange(`${scoreColLetter}${row}`).clear(Excel.ClearApplyTo.contents);
        }
      }
    }

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
  sheet.getRange('A4:P4').format.fill.clear();

  // Highlight current day's task and score header columns
  const taskCol = state.weekly.dIndex * 2 + 3;  // 3,5,7,9,11,13,15
  const scoreCol = state.weekly.dIndex * 2 + 4; // 4,6,8,10,12,14,16
  const taskColLetter = indexToColumnLetter(taskCol - 1);
  const scoreColLetter = indexToColumnLetter(scoreCol - 1);

  sheet.getRange(`${taskColLetter}4`).format.fill.color = CONFIG.COLORS.TODAY_HIGHLIGHT;
  sheet.getRange(`${scoreColLetter}4`).format.fill.color = CONFIG.COLORS.TODAY_HIGHLIGHT;

  await context.sync();
  console.log('Highlighted current day column:', state.weekly.dIndex);
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
    sheet.getRange('B5:B' + CONFIG.WEEKLY.lastTimeLine).format.fill.clear();

    // Get current time as decimal (e.g., 15.75 for 15:45)
    const now = new Date();
    const currentHour = now.getHours();
    const currentMinutes = now.getMinutes();
    const currentTimeDecimal = currentHour + currentMinutes / 60;

    console.log('=== HIGHLIGHT TIME DEBUG ===');
    console.log('Current time:', currentHour + ':' + currentMinutes, '= decimal:', currentTimeDecimal);
    console.log('Looking in rows', CONFIG.WEEKLY.DATA_START_ROW, 'to', CONFIG.WEEKLY.lastTimeLine);

    // Get time column values
    const timeRange = sheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.lastTimeLine}`);
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
      const taskCol = state.weekly.dIndex * 2 + 3;
      const scoreCol = state.weekly.dIndex * 2 + 4;
      const taskColLetter = indexToColumnLetter(taskCol - 1);
      const scoreColLetter = indexToColumnLetter(scoreCol - 1);

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
 * Refresh current time highlighting
 * Call this periodically or on demand
 */
async function refreshTimeHighlight() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.WEEKLY_SHEET);
      await context.sync();

      if (!sheet.isNullObject) {
        await highlightCurrentTimeRow(context, sheet);
      }
    });
    showStatus('Time highlight refreshed!', 'success');
  } catch (error) {
    showStatus('Error refreshing time: ' + error.message, 'error');
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
      row >= CONFIG.WEEKLY.DATA_START_ROW && row <= CONFIG.WEEKLY.lastTimeLine) {
    showStatus('Select a task from the dropdown, or use Add Task to create new.', 'info');
    return;
  }

  // Score column selection (even columns 4-16, rows 5+)
  if (CONFIG.WEEKLY.SCORE_COLUMNS.includes(colIndex) &&
      row >= CONFIG.WEEKLY.DATA_START_ROW && row < CONFIG.WEEKLY.scoreLine) {

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
    const tasksRange = tasksSheet.getRange(`A4:A${state.weekly.taskl}`);
    tasksRange.load('values');
    await context.sync();

    const tasks = tasksRange.values.filter(t => t[0] && t[0] !== '');
    if (tasks.length === 0) {
      showStatus('No tasks available for random pick', 'error');
      return;
    }

    // Current day task column
    const taskColumn = state.weekly.dIndex * 2 + 3; // 3,5,7,9,11,13,15
    const taskColLetter = indexToColumnLetter(taskColumn - 1);

    // Get time column and task column
    const timeRange = weeklySheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.lastTimeLine}`);
    const taskRange = weeklySheet.getRange(`${taskColLetter}${CONFIG.WEEKLY.DATA_START_ROW}:${taskColLetter}${CONFIG.WEEKLY.lastTimeLine}`);

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
 * Equivalent to VBA Worksheet_Change
 * @param {Excel.RequestContext} context - Excel context
 * @param {number} row - Row number
 * @param {number} col - Column number (1-based)
 * @param {number} newScore - The new score value
 */
async function processWeeklyScoreChange(context, row, col, newScore) {
  const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
  const tasksSheet = context.workbook.worksheets.getItem(CONFIG.TASKS_SHEET);

  // Get task name from adjacent cell
  const taskColLetter = indexToColumnLetter(col - 2);
  const taskCell = weeklySheet.getRange(`${taskColLetter}${row}`);
  taskCell.load('values');
  await context.sync();

  const taskName = taskCell.values[0][0];
  if (!taskName) return;

  // Calculate target day
  const dayIndex = Math.floor((col - 4) / 2);
  const targetDay = new Date(state.weekly.lastMonday);
  targetDay.setDate(state.weekly.lastMonday.getDate() + dayIndex);

  // Find task in Tasks sheet
  const tasksRange = tasksSheet.getRange(`A4:A${state.weekly.taskl}`);
  tasksRange.load('values');
  await context.sync();

  let taskIndex = -1;
  let othersIndex = -1;
  let taskWeight = 1;

  for (let i = 0; i < tasksRange.values.length; i++) {
    if (tasksRange.values[i][0] === taskName) {
      taskIndex = i + 4;
      const weightCell = tasksSheet.getRange(`B${taskIndex}`);
      weightCell.load('values');
      await context.sync();
      taskWeight = parseFloat(weightCell.values[0][0]) || 1;
      break;
    }
    if (tasksRange.values[i][0] === 'others') {
      othersIndex = i + 4;
    }
  }

  // If task not found, use "others"
  if (taskIndex === -1) {
    if (othersIndex !== -1) {
      taskIndex = othersIndex;
      const weightCell = tasksSheet.getRange(`B${taskIndex}`);
      weightCell.load('values');
      await context.sync();
      taskWeight = parseFloat(weightCell.values[0][0]) || 1;
    } else {
      // Create "others"
      taskIndex = state.weekly.taskl + 1;
      tasksSheet.getRange(`A${taskIndex}`).values = [['others']];
      tasksSheet.getRange(`B${taskIndex}`).values = [[1]];
      tasksSheet.getRange(`C${taskIndex}`).values = [[formatDateTime(new Date())]];
      state.weekly.taskl = taskIndex;
    }
  }

  // Calculate weighted score
  const weightedScore = taskWeight * newScore;

  // Apply colors to the individual score cell
  const scoreColLetter = indexToColumnLetter(col - 1);
  const scoreCell = weeklySheet.getRange(`${scoreColLetter}${row}`);

  let color;
  if (weightedScore > 0) {
    color = CONFIG.COLORS.POSITIVE;
  } else if (weightedScore < 0) {
    color = CONFIG.COLORS.NEGATIVE;
  } else {
    color = CONFIG.COLORS.NEUTRAL;
  }

  scoreCell.format.fill.color = color;
  taskCell.format.fill.color = color;

  // Update the daily total in scoreLine row (row 38)
  // The score column for this day is the same as the individual score column
  const dailyTotalCell = weeklySheet.getRange(`${scoreColLetter}${CONFIG.WEEKLY.scoreLine}`);
  dailyTotalCell.load('values');
  await context.sync();

  const currentDailyTotal = parseFloat(dailyTotalCell.values[0][0]) || 0;
  dailyTotalCell.values = [[currentDailyTotal + weightedScore]];

  // Update summary sheet
  await updateSummary(context, weightedScore > 0 ? weightedScore : 0, weightedScore < 0 ? weightedScore : 0);

  // Update task statistics in Tasks sheet
  tasksSheet.getRange(`D${taskIndex}`).values = [[formatDateTime(new Date())]];

  const countCell = tasksSheet.getRange(`F${taskIndex}`);
  const taskScoreCell = tasksSheet.getRange(`G${taskIndex}`);
  countCell.load('values');
  taskScoreCell.load('values');
  await context.sync();

  countCell.values = [[(parseInt(countCell.values[0][0]) || 0) + 1]];
  taskScoreCell.values = [[(parseFloat(taskScoreCell.values[0][0]) || 0) + weightedScore]];

  await context.sync();

  showStatus(`📝 "${taskName}" scored: ${weightedScore.toFixed(2)} (Daily: ${(currentDailyTotal + weightedScore).toFixed(2)})`, 'success');
}

/**
 * Update Summary sheet with scores
 * @param {Excel.RequestContext} context - Excel context
 * @param {number} positiveScore - Positive score to add
 * @param {number} negativeScore - Negative score to add
 */
async function updateSummary(context, positiveScore, negativeScore) {
  try {
    const summarySheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.SUMMARY_SHEET);
    await context.sync();

    if (summarySheet.isNullObject) return;

    const todayStr = formatDateYYYYMMDD(new Date());

    // Find or create today's row
    const summaryRange = summarySheet.getRange(`${CONFIG.SUMMARY.DATE_COLUMN}1:${CONFIG.SUMMARY.DATE_COLUMN}${state.weekly.summaryL + 1}`);
    summaryRange.load('values');
    await context.sync();

    let summaryRow = -1;
    for (let i = 0; i < summaryRange.values.length; i++) {
      if (String(summaryRange.values[i][0]) === todayStr) {
        summaryRow = i + 1;
        break;
      }
    }

    if (summaryRow === -1) {
      summaryRow = state.weekly.summaryL + 1;
      summarySheet.getRange(`${CONFIG.SUMMARY.DATE_COLUMN}${summaryRow}`).values = [[todayStr]];
      state.weekly.summaryL = summaryRow;
    }

    // Update positive score (Column D from config)
    if (positiveScore > 0) {
      const posCell = summarySheet.getRange(`${CONFIG.SUMMARY.POSITIVE_SCORE_COLUMN}${summaryRow}`);
      posCell.load('values');
      await context.sync();
      posCell.values = [[(parseFloat(posCell.values[0][0]) || 0) + positiveScore]];
    }

    // Update negative score (Column E from config)
    if (negativeScore < 0) {
      const negCell = summarySheet.getRange(`${CONFIG.SUMMARY.NEGATIVE_SCORE_COLUMN}${summaryRow}`);
      negCell.load('values');
      await context.sync();
      negCell.values = [[(parseFloat(negCell.values[0][0]) || 0) + negativeScore]];
    }

    // Update total score (Column F from config) = positive + negative
    const totalCell = summarySheet.getRange(`${CONFIG.SUMMARY.TOTAL_SCORE_COLUMN}${summaryRow}`);
    totalCell.load('values');
    await context.sync();
    const currentTotal = parseFloat(totalCell.values[0][0]) || 0;
    totalCell.values = [[currentTotal + positiveScore + negativeScore]];

    await context.sync();
  } catch (error) {
    console.error('Update summary error:', error);
  }
}

// ============================================================================
// ARCHIVE & NEW WEEK FUNCTIONS
// ============================================================================

/**
 * Archive the current week's data and start a new week
 * This exports data as CSV, then clears for new week
 */
async function archiveAndStartNewWeek() {
  try {
    showStatus('📦 Archiving week data...', 'info');

    // Step 1: Export current week data
    const weekData = await exportWeekData();

    if (weekData) {
      // Step 2: Download as CSV
      downloadCSV(weekData.csv, weekData.filename);

      // Step 3: Show instructions for OneDrive copy
      showArchiveInstructions();
    }

    // Step 4: Ask user to confirm before clearing
    // (In a real scenario, you'd use a dialog, but for simplicity we'll proceed)
    showStatus('📥 Week archived! Click "Start New Week" to clear data.', 'success');

  } catch (error) {
    console.error('Archive error:', error);
    showStatus('Error archiving: ' + error.message, 'error');
  }
}

/**
 * Export current week data to CSV format
 * @returns {Object} Object with csv string and filename
 */
async function exportWeekData() {
  try {
    let csvContent = '';
    let weekLabel = '';

    await Excel.run(async (context) => {
      const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

      // Get date info from B4
      const dateCell = weeklySheet.getRange('B4');
      dateCell.load('values');

      // Get day headers (D4:P4)
      const headerRange = weeklySheet.getRange('D4:P4');
      headerRange.load('values');

      // Get time column
      const timeRange = weeklySheet.getRange(`B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.lastTimeLine}`);
      timeRange.load('values');

      // Get all task and score data (C5:P[timel])
      const dataRange = weeklySheet.getRange(`C${CONFIG.WEEKLY.DATA_START_ROW}:P${CONFIG.WEEKLY.lastTimeLine}`);
      dataRange.load('values');

      await context.sync();

      // Build week label for filename
      const dateStr = String(dateCell.values[0][0] || '');
      const firstDay = headerRange.values[0][0];
      const lastDay = headerRange.values[0][headerRange.values[0].length - 1];
      weekLabel = `${dateStr.replace(' ', '-')}_${firstDay}-${lastDay}`;

      // Build CSV header
      const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
      let headers = ['Time'];
      for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
        headers.push(`${days[d]}_Task`);
        headers.push(`${days[d]}_Score`);
      }
      csvContent = headers.join(',') + '\n';

      // Build CSV rows
      for (let i = 0; i < timeRange.values.length; i++) {
        const time = timeRange.values[i][0];
        if (time === '' || time === null) continue;

        // Format time
        let timeStr;
        if (typeof time === 'number') {
          const hours = Math.floor(time * 24);
          const mins = Math.round((time * 24 - hours) * 60);
          timeStr = `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
        } else {
          timeStr = String(time);
        }

        let row = [escapeCSV(timeStr)];

        // Add task/score pairs for each day
        for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
          const taskCol = d * 2;      // 0, 2, 4, 6, 8, 10, 12
          const scoreCol = d * 2 + 1; // 1, 3, 5, 7, 9, 11, 13
          const task = dataRange.values[i][taskCol] || '';
          const score = dataRange.values[i][scoreCol];
          row.push(escapeCSV(String(task)));
          row.push(score !== null && score !== '' ? score : '');
        }

        csvContent += row.join(',') + '\n';
      }
    });

    const filename = `Weekly_${weekLabel}.csv`;
    return { csv: csvContent, filename: filename };

  } catch (error) {
    console.error('Export error:', error);
    showStatus('Error exporting: ' + error.message, 'error');
    return null;
  }
}

/**
 * Escape a value for CSV (handle commas, quotes, newlines)
 * @param {string} value - Value to escape
 * @returns {string} Escaped value
 */
function escapeCSV(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * Download a string as a CSV file
 * @param {string} content - CSV content
 * @param {string} filename - Filename for download
 */
function downloadCSV(content, filename) {
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  console.log('Downloaded:', filename);
}

/**
 * Show instructions for creating a copy in OneDrive
 */
function showArchiveInstructions() {
  const instructions = `
📁 To save a copy of the Excel file:

1. In Excel Online:
   • File → Save As → Save a Copy
   • Rename with week date

2. In OneDrive:
   • Right-click the file
   • Select "Copy to"
   • Rename the copy

3. Version History:
   • File → Info → Version History
   • Restore any previous version
  `;
  console.log(instructions);
}

/**
 * Start a new week (clear data and set new dates)
 * Call this after archiving
 */
async function startNewWeekFromUI() {
  try {
    showStatus('🗓️ Starting new week...', 'info');

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

      // Clear the weekly data
      await clearForNewWeek(context);

      // Set new dates
      await setNewWeekDates(context, sheet);

      // Highlight current day
      await highlightCurrentDay(context, sheet);

      // Highlight current time
      await highlightCurrentTimeRow(context, sheet);
    });

    showStatus('✅ New week started! Remember to save a copy for archive.', 'success');

  } catch (error) {
    console.error('Start new week error:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

/**
 * Export all important sheets as separate CSV files for backup
 * This exports Weekly, Goals, and Charter sheets
 */
async function exportWeeklyAsXLS() {
  try {
    showStatus('📊 Exporting all sheets...', 'info');

    let weekLabel = '';

    await Excel.run(async (context) => {
      const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

      // Get date info for filename
      const dateCell = weeklySheet.getRange('B4');
      dateCell.load('values');

      const headerRange = weeklySheet.getRange('D4:P4');
      headerRange.load('values');

      await context.sync();

      // Build week label for filename
      const dateStr = String(dateCell.values[0][0] || '').trim();
      if (dateStr && headerRange.values[0].length > 0) {
        const firstDay = headerRange.values[0][0];
        const lastDay = headerRange.values[0][headerRange.values[0].length - 1];
        weekLabel = `${dateStr.replace(' ', '-')}_${firstDay}-${lastDay}`;
      } else {
        weekLabel = formatDateYYYYMMDD(new Date());
      }
    });

    // Export each sheet
    const sheetsToExport = [
      { name: CONFIG.WEEKLY_SHEET, prefix: 'Weekly' },
      { name: 'Goals', prefix: 'Goals' },
      { name: 'Charter', prefix: 'Charter' }
    ];

    let exportedCount = 0;

    for (const sheetInfo of sheetsToExport) {
      const csvContent = await exportSheetAsCSV(sheetInfo.name);
      if (csvContent) {
        const filename = `${sheetInfo.prefix}_${weekLabel}.csv`;
        downloadCSV(csvContent, filename);
        exportedCount++;
        // Small delay between downloads to prevent browser blocking
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }

    if (exportedCount > 0) {
      showStatus(`📊 Exported ${exportedCount} sheets as CSV files!`, 'success');
    } else {
      showStatus('No sheets found to export.', 'warning');
    }

  } catch (error) {
    console.error('Export error:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

/**
 * Export a single sheet as CSV
 * @param {string} sheetName - Name of the sheet to export
 * @returns {string|null} CSV content or null if sheet not found
 */
async function exportSheetAsCSV(sheetName) {
  try {
    let csvContent = '';

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();

      if (sheet.isNullObject) {
        console.log(`Sheet "${sheetName}" not found, skipping...`);
        return;
      }

      // Get the used range to export entire sheet layout
      const usedRange = sheet.getUsedRange();
      usedRange.load('values');

      await context.sync();

      // Convert entire used range to CSV (preserving exact layout)
      for (const row of usedRange.values) {
        csvContent += row.map(cell => {
          // Format time values properly
          if (typeof cell === 'number' && cell > 0 && cell < 1) {
            // This is likely a time value (Excel stores time as fraction of day)
            const hours = Math.floor(cell * 24);
            const mins = Math.round((cell * 24 - hours) * 60);
            return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
          }
          return escapeCSV(cell);
        }).join(',') + '\n';
      }
    });

    return csvContent || null;

  } catch (error) {
    console.error(`Error exporting sheet "${sheetName}":`, error);
    return null;
  }
}

/**
 * Download base64 content as an XLSX file
 * @param {string} base64Content - Base64 encoded workbook content
 * @param {string} filename - Filename for download
 */
function downloadXLSX(base64Content, filename) {
  // Convert base64 to binary
  const byteCharacters = atob(base64Content);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);

  // Create blob and download
  const blob = new Blob([byteArray], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  console.log('Downloaded:', filename);
}

/**
 * Export Summary sheet data as CSV
 */
async function exportSummaryData() {
  try {
    showStatus('📊 Exporting summary...', 'info');

    let csvContent = '';

    await Excel.run(async (context) => {
      const summarySheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.SUMMARY_SHEET);
      await context.sync();

      if (summarySheet.isNullObject) {
        showStatus('Summary sheet not found!', 'error');
        return;
      }

      const usedRange = summarySheet.getUsedRange();
      usedRange.load('values');
      await context.sync();

      // Convert to CSV
      for (const row of usedRange.values) {
        csvContent += row.map(cell => escapeCSV(cell)).join(',') + '\n';
      }
    });

    if (csvContent) {
      const today = formatDateYYYYMMDD(new Date());
      downloadCSV(csvContent, `Summary_${today}.csv`);
      showStatus('📊 Summary exported!', 'success');
    }

  } catch (error) {
    console.error('Export summary error:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

// Export for use in other modules
window.initializeWeeklySheet = initializeWeeklySheet;
window.initializeWeeklyOnOpen = initializeWeeklyOnOpen;
window.initializeTasksSheet = initializeTasksSheet;
window.initializeSummarySheet = initializeSummarySheet;
window.setNewWeekDates = setNewWeekDates;
window.clearForNewWeek = clearForNewWeek;
window.highlightCurrentDay = highlightCurrentDay;
window.highlightCurrentTimeRow = highlightCurrentTimeRow;
window.refreshTimeHighlight = refreshTimeHighlight;
window.handleWeeklySelection = handleWeeklySelection;
window.randomPick = randomPick;
window.processWeeklyScoreChange = processWeeklyScoreChange;
window.updateSummary = updateSummary;
window.archiveAndStartNewWeek = archiveAndStartNewWeek;
window.exportWeekData = exportWeekData;
window.startNewWeekFromUI = startNewWeekFromUI;
window.exportSummaryData = exportSummaryData;
window.exportWeeklyAsXLS = exportWeeklyAsXLS;
window.downloadCSV = downloadCSV;
window.downloadXLSX = downloadXLSX;
