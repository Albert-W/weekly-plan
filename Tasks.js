/**
 * Tasks-sheet domain logic (Google Sheets edition).
 *
 * Ported from src/taskpane/js/tasks.js. The Office build tracked the
 * last task row in an in-memory state object; GAS is stateless per
 * execution, so getLastTaskRow_() computes it from the sheet each call.
 */

/**
 * 1-based index of the last task row (or DATA_START_ROW - 1 when empty).
 * Replaces state.weekly.lastTaskRow.
 * @returns {number}
 */
function getLastTaskRow_() {
  const sheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!sheet) return CONFIG.TASKS.DATA_START_ROW - 1;
  return Math.max(getLastRowInColumn_(sheet, 1), CONFIG.TASKS.DATA_START_ROW - 1);
}

/**
 * Create a new task: appends name (A), weight (B), created timestamp (C).
 * @param {string} name
 * @param {number} weight defaults to 1
 * @returns {{row: number, name: string, weight: number}}
 */
function createTask(name, weight) {
  const trimmedName = String(name || '').trim();
  if (!trimmedName) {
    throw new Error('Task name is required');
  }
  const numericWeight = isFinite(weight) ? Number(weight) : 1;

  const sheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!sheet) {
    throw new Error('Tasks sheet not found. Run "Set up sheets" first.');
  }

  const newRow = Math.max(getLastRowInColumn_(sheet, 1) + 1, CONFIG.TASKS.DATA_START_ROW);
  sheet.getRange('A' + newRow).setValue(trimmedName);
  sheet.getRange('B' + newRow).setValue(numericWeight);
  sheet.getRange('C' + newRow).setValue(formatDateTime(new Date()));

  // Keep the Weekly task-cell dropdowns in sync (now includes habits too).
  refreshDropdownList_();

  return { row: newRow, name: trimmedName, weight: numericWeight };
}

/**
 * Delete a task by name (first matching row in column A). Returns true
 * if a row was removed.
 * @param {string} name
 * @returns {boolean}
 */
function deleteTask(name) {
  const trimmed = String(name || '').trim();
  if (!trimmed) return false;
  const sheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!sheet) return false;

  const lastRow = getLastRowInColumn_(sheet, 1);
  if (lastRow < CONFIG.TASKS.DATA_START_ROW) return false;

  const start = CONFIG.TASKS.DATA_START_ROW;
  const names = sheet.getRange(start, 1, lastRow - start + 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === trimmed) {
      sheet.deleteRow(start + i);
      refreshDropdownList_();
      return true;
    }
  }
  return false;
}

/**
 * Sort task rows by weight (column B) descending.
 * Mirrors sortHabits(): keeps each task's full row (A–G) together.
 * @returns {string} status message
 */
function sortTasks() {
  const sheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!sheet) {
    toast_('Tasks sheet not found.', 'Weekly Plan');
    return 'Tasks sheet not found.';
  }
  const start = CONFIG.TASKS.DATA_START_ROW;
  const lastRow = getLastTaskRow_();
  if (lastRow < start) return 'No tasks to sort.';

  const weightCol = columnLetterToIndex('B') + 1; // B -> 2
  const lastCol = columnLetterToIndex('G') + 1; // G -> 7 (Total Score)
  sheet
    .getRange(start, 1, lastRow - start + 1, lastCol)
    .sort({ column: weightCol, ascending: false });

  toast_('Tasks sorted by weight.', 'Weekly Plan');
  return 'Tasks sorted by weight.';
}
