/**
 * Weekly/Timetable sheet logic (Google Sheets edition).
 *
 * Ported from src/taskpane/js/weekly.js. The async Excel.run/context.sync
 * model becomes synchronous SpreadsheetApp calls. Score-cell colors are
 * now handled by conditional formatting (see Setup.gs), so the score
 * handler only colors the matching task cell. The 60s client ticker
 * lives in the sidebar instead of here. currentDayIndex / lastMonday
 * are recomputed each call (GAS is stateless).
 */

/**
 * Reconstruct the Monday (00:00) the Weekly sheet currently represents,
 * from B4 + the first day-number header (D4). B4 is normally the string
 * "yyyy mm", but Google Sheets may auto-convert that entry into a real
 * Date, so both forms are handled. Null only when truly unset.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Date|null}
 */
function getSheetWeekMonday_(sheet) {
  const W = CONFIG.WEEKLY;
  const dateVal = sheet.getRange(W.DATE_CELL).getValue();
  const firstDay = parseInt(sheet.getRange(W.FIRST_DAY_HEADER_CELL).getValue(), 10) || 0;
  if (!firstDay) return null;

  let year;
  let month; // 0-based
  if (dateVal instanceof Date) {
    year = dateVal.getFullYear();
    month = dateVal.getMonth();
  } else {
    const parts = String(dateVal || '').trim().split(/\s+/);
    if (parts.length < 2) return null;
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10) - 1;
  }
  if (isNaN(year) || isNaN(month)) return null;

  const d = new Date(year, month, firstDay);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Initialize the Weekly sheet (run on open / new-day).
 * Detects a new week (>= 7 days since the sheet's Monday) and rolls
 * over, otherwise just refreshes highlights.
 */
function initializeWeeklyOnOpen() {
  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) return;

  const sheetLastMonday = getSheetWeekMonday_(sheet);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (sheetLastMonday) {
    const diffDays = daysBetween(sheetLastMonday, today);
    if (diffDays >= 7) {
      toast_('New week detected — archiving and resetting…', 'Weekly Plan');
      doWeekRollover_(true);
      return;
    }
  } else {
    setNewWeekDates(sheet);
  }

  highlightCurrentDay(sheet);
  highlightCurrentTimeRow(sheet);
}

/**
 * Write the current week's dates: B4 = "yyyy mm" and day numbers in the
 * score-column headers (D4, F4, … P4).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setNewWeekDates(sheet) {
  const W = CONFIG.WEEKLY;
  const newMonday = getMonday(new Date());

  const yearMonth =
    newMonday.getFullYear() + ' ' + String(newMonday.getMonth() + 1).padStart(2, '0');
  sheet.getRange(W.DATE_CELL).setValue(yearMonth);

  for (let i = 0; i < W.DAYS_IN_WEEK; i++) {
    const dayDate = new Date(newMonday);
    dayDate.setDate(newMonday.getDate() + i);
    const colLetter = getScoreColLetterForDay(i);
    sheet.getRange(colLetter + W.HEADER_ROW).setValue(dayDate.getDate());
  }
}

/**
 * Clear the grid for a new week. Only clears task+score pairs for rows
 * that actually had a score recorded; task-only rows are preserved
 * (matches the original VBA / Office behavior).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearForNewWeek(sheet) {
  const W = CONFIG.WEEKLY;
  const dataStart = W.DATA_START_ROW;
  const dataEnd = W.SCORE_ROW - 1;
  const firstCol = getTaskColForDay(0); // C (col 3)
  const colSpan = getScoreColForDay(W.DAYS_IN_WEEK - 1) - firstCol + 1; // C..P = 14

  // Clear any manual task-cell backgrounds and the totals row.
  sheet
    .getRange(dataStart, firstCol, W.SCORE_ROW - dataStart + 1, colSpan)
    .setBackground(null);
  sheet.getRange(W.SCORE_ROW, firstCol, 1, colSpan).clearContent();

  const values = sheet
    .getRange(dataStart, firstCol, dataEnd - dataStart + 1, colSpan)
    .getValues();

  for (let day = 0; day < W.DAYS_IN_WEEK; day++) {
    const taskColOffset = day * 2;
    const scoreColOffset = day * 2 + 1;
    for (let i = 0; i < values.length; i++) {
      const scoreVal = values[i][scoreColOffset];
      if (scoreVal !== '' && scoreVal !== null) {
        const row = dataStart + i;
        sheet.getRange(row, firstCol + taskColOffset).clearContent();
        sheet.getRange(row, firstCol + scoreColOffset).clearContent();
      }
    }
  }
}

/**
 * Highlight the current day's header cells (row 4).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function highlightCurrentDay(sheet) {
  const W = CONFIG.WEEKLY;
  sheet.getRange(W.HEADER_ROW_RANGE).setBackground(null);

  const dayIndex = getCurrentDayIndex();
  const taskColLetter = getTaskColLetterForDay(dayIndex);
  const scoreColLetter = getScoreColLetterForDay(dayIndex);
  sheet.getRange(taskColLetter + W.HEADER_ROW).setBackground(CONFIG.COLORS.TODAY_HIGHLIGHT);
  sheet.getRange(scoreColLetter + W.HEADER_ROW).setBackground(CONFIG.COLORS.TODAY_HIGHLIGHT);
}

/**
 * 1-based Weekly row whose time value best matches the current clock time
 * (the latest slot at/just before now), or -1 if none.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getCurrentTimeRow_(sheet) {
  const W = CONFIG.WEEKLY;
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;

  const now = new Date();
  const currentTimeDecimal = now.getHours() + now.getMinutes() / 60;

  const timeValues = sheet
    .getRange(W.DATA_START_ROW, W.TIME_COLUMN, slotCount, 1)
    .getValues();

  let bestRowIndex = -1;
  let bestTimeValue = -1;
  for (let i = 0; i < timeValues.length; i++) {
    const cellTime = timeValues[i][0];
    if (cellTime === '' || cellTime === null || cellTime === undefined) continue;

    let timeValue = null;
    if (cellTime instanceof Date) {
      timeValue = cellTime.getHours() + cellTime.getMinutes() / 60;
    } else if (typeof cellTime === 'number') {
      if (cellTime >= 0 && cellTime <= 1) timeValue = cellTime * 24;
      else if (cellTime <= 24) timeValue = cellTime;
      else continue;
    } else if (typeof cellTime === 'string') {
      const m = cellTime.match(/^(\d{1,2}):(\d{2})$/);
      if (m) timeValue = parseInt(m[1], 10) + parseInt(m[2], 10) / 60;
      else continue;
    } else {
      continue;
    }

    if (timeValue <= currentTimeDecimal + 0.1 && timeValue > bestTimeValue) {
      bestTimeValue = timeValue;
      bestRowIndex = i;
    }
  }

  return bestRowIndex < 0 ? -1 : W.DATA_START_ROW + bestRowIndex;
}

/**
 * Highlight the time row matching the current clock time.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function highlightCurrentTimeRow(sheet) {
  const W = CONFIG.WEEKLY;
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;

  // Clear previous time-column highlight.
  sheet.getRange(W.DATA_START_ROW, W.TIME_COLUMN, slotCount, 1).setBackground(null);

  const row = getCurrentTimeRow_(sheet);
  if (row < 0) return;
  sheet.getRange(row, W.TIME_COLUMN).setBackground(CONFIG.COLORS.CURRENT_TIME);

  // If the current day's slot has no score yet, highlight task+score too.
  const dayIndex = getCurrentDayIndex();
  const taskCol = getTaskColForDay(dayIndex);
  const scoreCol = getScoreColForDay(dayIndex);
  const scoreVal = sheet.getRange(row, scoreCol).getValue();
  if (scoreVal === '' || scoreVal === null) {
    sheet.getRange(row, taskCol).setBackground(CONFIG.COLORS.CURRENT_TIME);
    sheet.getRange(row, scoreCol).setBackground(CONFIG.COLORS.CURRENT_TIME);
  }
}

/**
 * Fill empty current-day task slots (that have a time) with random tasks.
 * @returns {number} number of slots filled
 */
function randomPick() {
  const weeklySheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  const tasksSheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!weeklySheet || !tasksSheet) {
    toast_('Weekly or Tasks sheet not found.', 'Weekly Plan');
    return 0;
  }
  const W = CONFIG.WEEKLY;

  const lastTaskRow = getLastTaskRow_();
  const tasks = [];
  if (lastTaskRow >= CONFIG.TASKS.DATA_START_ROW) {
    const taskVals = tasksSheet
      .getRange(CONFIG.TASKS.DATA_START_ROW, 1, lastTaskRow - CONFIG.TASKS.DATA_START_ROW + 1, 1)
      .getValues();
    for (let i = 0; i < taskVals.length; i++) {
      if (taskVals[i][0] && taskVals[i][0] !== '') tasks.push(taskVals[i][0]);
    }
  }
  if (tasks.length === 0) {
    toast_('No tasks available for random pick.', 'Weekly Plan');
    return 0;
  }

  const dayIndex = getCurrentDayIndex();
  const taskCol = getTaskColForDay(dayIndex);
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;
  const timeVals = weeklySheet.getRange(W.DATA_START_ROW, W.TIME_COLUMN, slotCount, 1).getValues();
  const taskVals = weeklySheet.getRange(W.DATA_START_ROW, taskCol, slotCount, 1).getValues();

  let filled = 0;
  for (let i = 0; i < slotCount; i++) {
    const hasTime = timeVals[i][0] !== '' && timeVals[i][0] !== null;
    const hasTask = taskVals[i][0] !== '' && taskVals[i][0] !== null;
    if (hasTime && !hasTask) {
      const randomTask = tasks[Math.floor(Math.random() * tasks.length)];
      weeklySheet.getRange(W.DATA_START_ROW + i, taskCol).setValue(randomTask);
      filled++;
    }
  }

  toast_(
    filled > 0 ? 'Filled ' + filled + ' slot(s) with random tasks.' : 'No empty slots to fill.',
    'Weekly Plan'
  );
  return filled;
}

/**
 * Process a score entered in a Weekly score cell. Updates the task-cell
 * color, the daily total, the Tasks stats, and the Summary sheet.
 * @param {number} row 1-based row
 * @param {number} col 1-based score-column index
 * @param {number} newScore
 */
function processWeeklyScoreChange(row, col, newScore) {
  const weeklySheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  const tasksSheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!weeklySheet || !tasksSheet) return;

  const W = CONFIG.WEEKLY;
  const taskCol = col - 1; // 1-based task column (score col - 1)
  const lastTaskRow = getLastTaskRow_();

  // Mutable state set inside the lock, read afterwards for side effects.
  let weightedScore = 0;
  let newDailyTotal = 0;
  let taskName = '';
  let comboDays = 0;
  let isQuestTask = false;
  let isHabit = false;
  let habitRow = -1;

  // ------------------------------------------------------------------
  // Read-modify-write under the document lock. The quest combo is
  // read AND advanced atomically inside this lock (via
  // applyComboLocked_) so a concurrent habit/score handler can't read
  // a stale combo state and apply the wrong multiplier.
  // ------------------------------------------------------------------
  let finished = false;
  try {
    withLock_(function () {
      taskName = weeklySheet.getRange(row, taskCol).getValue();
      if (!taskName) return;

      const dailyTotalCell = weeklySheet.getRange(W.SCORE_ROW, col);
      const currentDailyTotal = parseFloat(dailyTotalCell.getValue()) || 0;

      // ------------------------------------------------------------------
      // Resolve the weight for this task / habit name.
      // 1) Habits sheet → weight = weight column (col C)
      // 2) Tasks sheet   → weight = task weight (col B)
      // 3) Fallback      → "others" row or auto-create (weight = 1)
      // ------------------------------------------------------------------
      let taskWeight = 1;
      isHabit = false;
      habitRow = -1;
      let taskRow = -1;
      let isNewTask = false;
      let currentCount = 0;
      let currentTaskScore = 0;

      // -- 1) Habits lookup --
      const habitsSheet = getSheetByName_(CONFIG.HABITS_SHEET);
      if (habitsSheet) {
        const lastHabitRow = getLastHabitRow_();
        const H = CONFIG.HABITS;
        if (lastHabitRow >= H.DATA_START_ROW) {
          const nameCol = columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1; // A → 1
          const habitData = habitsSheet
            .getRange(H.DATA_START_ROW, nameCol, lastHabitRow - H.DATA_START_ROW + 1, 3)
            .getValues();
          for (let i = 0; i < habitData.length; i++) {
            if (String(habitData[i][0]).trim() === taskName) {
              taskWeight = parseFloat(habitData[i][2]) || 1; // col C = index 2
              isHabit = true;
              habitRow = H.DATA_START_ROW + i;
              break;
            }
          }
        }
      }

      if (!isHabit) {
        // -- 2) Tasks lookup (existing logic) --
        const startRow = CONFIG.TASKS.DATA_START_ROW;
        let names = [];
        let weights = [];
        let counts = [];
        let scores = [];
        if (lastTaskRow >= startRow) {
          names = tasksSheet.getRange('A' + startRow + ':A' + lastTaskRow).getValues();
          weights = tasksSheet.getRange('B' + startRow + ':B' + lastTaskRow).getValues();
          counts = tasksSheet.getRange('F' + startRow + ':F' + lastTaskRow).getValues();
          scores = tasksSheet.getRange('G' + startRow + ':G' + lastTaskRow).getValues();
        }

        let othersRow = -1;
        for (let i = 0; i < names.length; i++) {
          const nm = names[i][0];
          if (nm === taskName) {
            taskRow = i + startRow;
            break;
          }
          if (nm === CONFIG.TASKS.FALLBACK_NAME) othersRow = i + startRow;
        }

        let lookupIndex;
        if (taskRow !== -1) {
          lookupIndex = taskRow - startRow;
        } else if (othersRow !== -1) {
          taskRow = othersRow;
          lookupIndex = othersRow - startRow;
        } else {
          taskRow = lastTaskRow + 1;
          lookupIndex = -1;
          isNewTask = true;
        }

        taskWeight = isNewTask ? 1 : parseFloat(weights[lookupIndex][0]) || 1;
        currentCount = isNewTask ? 0 : parseInt(counts[lookupIndex][0], 10) || 0;
        currentTaskScore = isNewTask ? 0 : parseFloat(scores[lookupIndex][0]) || 0;
      }

      weightedScore = taskWeight * newScore;
      if (weightedScore > 0) {
        // Daily Quest bonus — use the right lookup for habits vs tasks.
        const qMult = isHabit
          ? questHabitMultiplier_(taskName)
          : questTaskMultiplier_(taskName);
        if (qMult > 1) {
          // Atomically read AND advance the combo inside this lock
          // so the multiplier matches the state we persist.
          const combo = applyComboLocked_();
          comboDays = combo.days;
          weightedScore *= qMult * combo.multiplier;
          isQuestTask = true;
        } else {
          weightedScore *= qMult;
        }
      }
      newDailyTotal = currentDailyTotal + weightedScore;

      const color =
        weightedScore > 0
          ? CONFIG.COLORS.POSITIVE
          : weightedScore < 0
          ? CONFIG.COLORS.NEGATIVE
          : CONFIG.COLORS.NEUTRAL;
      const now = formatDateTime(new Date());

      // Color the task cell to match (score cell is colored by conditional formatting).
      weeklySheet.getRange(row, taskCol).setBackground(color);
      // Normalize the score cell to a real number so the numeric
      // conditional-format rules apply (dropdown picks can land as text).
      weeklySheet.getRange(row, col).setValue(newScore);
      dailyTotalCell.setValue(newDailyTotal);

      // Habits don't have rows in the Tasks sheet — skip the stats update.
      if (!isHabit) {
        if (isNewTask) {
          tasksSheet.getRange('A' + taskRow).setValue(CONFIG.TASKS.FALLBACK_NAME);
          tasksSheet.getRange('B' + taskRow).setValue(1);
          tasksSheet.getRange('C' + taskRow).setValue(now);
          tasksSheet.getRange('D' + taskRow).setValue(now);
          tasksSheet.getRange('F' + taskRow).setValue(1);
          tasksSheet.getRange('G' + taskRow).setValue(weightedScore);
        } else {
          tasksSheet.getRange('D' + taskRow).setValue(now);
          tasksSheet.getRange('F' + taskRow).setValue(currentCount + 1);
          tasksSheet.getRange('G' + taskRow).setValue(currentTaskScore + weightedScore);
        }
      }
      finished = true;
    });
  } catch (e) {
    Logger.log('processWeeklyScoreChange: ' + (e && e.message ? e.message : e));
    return;
  }

  if (!finished || !taskName) return;

  // Flush AFTER releasing the lock so we don't hold it during I/O.
  SpreadsheetApp.flush();

  // If a habit was scored positively, mark it done in the Habits sheet
  // (mirrors the checkbox path without double-counting summary/XP/boss).
  if (isHabit && weightedScore > 0 && habitRow > 0) {
    const hSheet = getSheetByName_(CONFIG.HABITS_SHEET);
    if (hSheet) {
      const dayIdx = findHabitsDayIndex(hSheet);
      if (dayIdx >= 0) {
        const H = CONFIG.HABITS;
        const dayCol = columnLetterToIndex(H.COLUMNS.DAY_START) + 1 + dayIdx;
        const curVal = parseInt(hSheet.getRange(habitRow, dayCol).getValue(), 10) || 0;
        hSheet.getRange(habitRow, dayCol).setValue(curVal + 1);
        const totalCell = hSheet.getRange(H.COLUMNS.TOTAL_COUNT + habitRow);
        const curTotal = parseInt(totalCell.getValue(), 10) || 0;
        totalCell.setValue(curTotal + 1);
        hSheet.getRange(H.COLUMNS.HABIT_NAME + habitRow).setBackground(CONFIG.COLORS.POSITIVE);
      }
    }
  }

  // Summary update owns its own lock — call after releasing ours.
  updateSummary(weightedScore > 0 ? weightedScore : 0, weightedScore < 0 ? weightedScore : 0);

  if (weightedScore > 0) {
    awardXp_(weightedScore);
    maybeAwardEarlyBird_(new Date());
    checkBossDefeat_();
  }

  // Mark quest done (acquires its own lock via tryWithLock_).
  let comboNote = '';
  if (isQuestTask) {
    markQuestDone_(isHabit ? 'habit' : 'task', taskName);
    if (comboDays > 1) comboNote = ' 🔥' + comboDays + 'd combo';
  }

  toast_(
    '"' + taskName + '" scored ' + weightedScore.toFixed(2) + ' (daily ' + newDailyTotal.toFixed(2) + ')' +
      (isQuestTask ? ' ⭐ Quest bonus!' : '') + comboNote,
    'Weekly Plan'
  );
}

/**
 * Core week rollover: optionally archive, clear, set new dates, highlight.
 * @param {boolean} archive
 * @returns {{url: string|null, rows: number}}
 */
function doWeekRollover_(archive) {
  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) return { url: null, rows: 0 };

  let result = { url: null, rows: 0 };
  if (archive) {
    result = safeInit_('Archive week failed', function () {
      return archiveWeek_();
    }) || { url: null, rows: 0 };
  }
  clearForNewWeek(sheet);
  setNewWeekDates(sheet);
  highlightCurrentDay(sheet);
  highlightCurrentTimeRow(sheet);
  return result;
}

/**
 * Re-derive the daily Scores row (row 38) and task-cell colors from the
 * current grid, in one pass. Use after editing scores on the mobile app,
 * where the live onEdit trigger doesn't fire.
 *
 * Deterministic & safe to re-run: daily totals and colors are fully
 * derived from the grid. It intentionally does NOT touch the Summary
 * sheet or Tasks cumulative stats (those aggregate across history and
 * also include habit completions, so re-deriving them from the weekly
 * grid alone would double-count / clobber habit data).
 *
 * LIMITATION: Scores are recomputed as plain `taskWeight × score` —
 * Daily Quest bonuses and streak-combo multipliers are NOT applied.
 * These depend on DocumentProperties state (which quest was active,
 * what the combo count was at the moment of scoring) and cannot be
 * reconstructed from the grid alone. After a recalculate, the Scores
 * row may differ from the original live-scored values.
 *
 * @returns {string} status message
 */
function recalculateWeek() {
  const weeklySheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!weeklySheet) {
    toast_('Weekly sheet not found.', 'Weekly Plan');
    return 'Weekly sheet not found.';
  }
  const W = CONFIG.WEEKLY;

  // Build a task name -> weight lookup from the Tasks sheet.
  const weightByName = {};
  const tasksSheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (tasksSheet) {
    const lastTaskRow = getLastTaskRow_();
    const start = CONFIG.TASKS.DATA_START_ROW;
    if (lastTaskRow >= start) {
      const names = tasksSheet.getRange(start, 1, lastTaskRow - start + 1, 1).getValues();
      const wts = tasksSheet.getRange(start, 2, lastTaskRow - start + 1, 1).getValues();
      for (let i = 0; i < names.length; i++) {
        if (names[i][0]) weightByName[names[i][0]] = parseFloat(wts[i][0]) || 1;
      }
    }
  }

  // Also add habits (name → weight, col C).  Habit weights are only
  // added when there is no task with the same name, so task weights take
  // precedence if a name exists in both sheets.
  const habitsSheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (habitsSheet) {
    const lastHabitRow = getLastHabitRow_();
    const H = CONFIG.HABITS;
    if (lastHabitRow >= H.DATA_START_ROW) {
      const nameCol = columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1;
      const habitData = habitsSheet
        .getRange(H.DATA_START_ROW, nameCol, lastHabitRow - H.DATA_START_ROW + 1, 3)
        .getValues();
      for (let i = 0; i < habitData.length; i++) {
        const hName = String(habitData[i][0]).trim();
        if (hName && weightByName[hName] === undefined) {
          weightByName[hName] = parseFloat(habitData[i][2]) || 1;
        }
      }
    }
  }

  const firstCol = getTaskColForDay(0); // C
  const colSpan = getScoreColForDay(W.DAYS_IN_WEEK - 1) - firstCol + 1; // C..P = 14
  const dataRows = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;
  const range = weeklySheet.getRange(W.DATA_START_ROW, firstCol, dataRows, colSpan);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  const dailyTotals = new Array(W.DAYS_IN_WEEK).fill(0);

  for (let day = 0; day < W.DAYS_IN_WEEK; day++) {
    const taskOff = day * 2;
    const scoreOff = day * 2 + 1;
    for (let i = 0; i < values.length; i++) {
      const rawScore = values[i][scoreOff];
      if (rawScore === '' || rawScore === null) {
        backgrounds[i][taskOff] = CONFIG.COLORS.CLEAR; // drop stale color
        continue;
      }
      const score = parseFloat(rawScore);
      if (isNaN(score)) continue;
      const taskName = values[i][taskOff];
      const weight =
        taskName && weightByName[taskName] !== undefined ? weightByName[taskName] : 1;
      const weighted = weight * score;
      dailyTotals[day] += weighted;
      backgrounds[i][taskOff] =
        weighted > 0
          ? CONFIG.COLORS.POSITIVE
          : weighted < 0
          ? CONFIG.COLORS.NEGATIVE
          : CONFIG.COLORS.NEUTRAL;
    }
  }

  range.setBackgrounds(backgrounds);
  for (let day = 0; day < W.DAYS_IN_WEEK; day++) {
    weeklySheet.getRange(W.SCORE_ROW, getScoreColForDay(day)).setValue(dailyTotals[day]);
  }

  toast_('Recalculated daily scores from the grid.', 'Weekly Plan');
  return 'Recalculated daily scores from the grid.';
}
