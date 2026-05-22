/**
 * Tasks-sheet domain logic for the Combined Tracker Add-in.
 *
 * Owns the Excel-side mutations on the Tasks sheet so other modules
 * (UI form handlers, weekly score processing, etc.) only have to
 * speak in terms of "create a task" / "update task stats" rather
 * than poking Range objects.
 */

/**
 * Create a new task in the Tasks sheet.
 *
 * Appends a row at the bottom of column A with:
 *   A: name
 *   B: weight (numeric multiplier)
 *   C: creation timestamp
 *
 * Updates state.weekly.taskl to the new last-row index.
 *
 * @param {Excel.RequestContext} context - Excel context owned by caller
 * @param {string} name - Task name (must be non-empty)
 * @param {number} weight - Score multiplier (defaults to 1)
 * @returns {Promise<{row: number, name: string, weight: number}>}
 */
async function createTask(context, name, weight) {
  const trimmedName = String(name || '').trim();
  if (!trimmedName) {
    throw new Error('Task name is required');
  }
  const numericWeight = Number.isFinite(weight) ? weight : 1;

  const tasksSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.TASKS_SHEET);
  await context.sync();
  if (tasksSheet.isNullObject) {
    throw new Error('Tasks sheet not found');
  }

  // Find the next available row by inspecting the used range of column A.
  const usedRange = tasksSheet.getRange('A:A').getUsedRange();
  usedRange.load('rowCount');
  await context.sync();

  const newRow = usedRange.rowCount + 1;

  tasksSheet.getRange(`A${newRow}`).values = [[trimmedName]];
  tasksSheet.getRange(`B${newRow}`).values = [[numericWeight]];
  tasksSheet.getRange(`C${newRow}`).values = [[formatDateTime(new Date())]];

  await context.sync();

  // Keep the in-memory last-row tracker in sync so the next
  // randomPick / score-change call sees the row.
  state.weekly.taskl = newRow;

  return { row: newRow, name: trimmedName, weight: numericWeight };
}

// Export for use in other modules
window.createTask = createTask;
