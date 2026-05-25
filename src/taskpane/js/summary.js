/**
 * Summary-sheet domain logic for the Combined Tracker Add-in.
 *
 * Init, accumulate-by-day update, and standalone CSV export.
 * Other modules (habits.js, weekly.js) call updateSummary as the
 * single entry point for adding positive/negative scores.
 */

/**
 * Initialize Summary sheet data
 * @param {Excel.RequestContext} context - Excel context
 */
async function initializeSummarySheet(context) {
  const sheet = context.workbook.worksheets.getItem(CONFIG.SUMMARY_SHEET);
  const usedRange = sheet.getRange('A:A').getUsedRange();
  usedRange.load('rowCount');
  await context.sync();

  state.weekly.lastSummaryRow = usedRange.rowCount;
  console.log('Summary sheet initialized, lastSummaryRow:', state.weekly.lastSummaryRow);
}

/**
 * Set new week dates in the Weekly sheet
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Weekly sheet
 */

/**
 * Update the Summary sheet with today's score deltas.
 *
 * Two-sync pattern (task #37):
 *   Sync 1 — load A1:F{lastSummaryRow+1} in one read. Both the
 *            date column (for find-or-create-today) AND the existing
 *            D/E/F values of every row come back in one shot.
 *   Sync 2 — queue all writes (date if new row, then D/E/F) and
 *            flush in one batch.
 *
 * The outer try/catch swallows ItemNotFound from getItem so the
 * function remains a silent no-op when the Summary sheet is absent
 * (preserving the previous getItemOrNullObject behavior with one
 * fewer sync).
 *
 * @param {Excel.RequestContext} context
 * @param {number} positiveScore - Score to accumulate into col D
 *   (only applied when > 0).
 * @param {number} negativeScore - Score to accumulate into col E
 *   (only applied when < 0).
 */
async function updateSummary(context, positiveScore, negativeScore) {
  try {
    // getItem throws if the sheet is absent. We let the outer catch
    // handle that — silent no-op, no sync needed for existence check.
    const summarySheet = context.workbook.worksheets.getItem(CONFIG.SUMMARY_SHEET);

    // Read date column + D/E/F for every existing row plus the next
    // candidate row in a single batch. One sync. (Previously this
    // function did 4-5 separate syncs — see CHANGELOG / task #37.)
    const lastRow = state.weekly.lastSummaryRow + 1;
    const allRange = summarySheet.getRange(`A1:F${lastRow}`);
    allRange.load('values');
    await context.sync();

    const todayStr = formatDateYYYYMMDD(new Date());
    const rows = allRange.values; // 2D array, [rowIdx][colIdx 0..5]

    // Find today's row, or fall through to "append".
    // Columns: A=0 (date), B=1, C=2, D=3 (pos), E=4 (neg), F=5 (total).
    let summaryRow = -1;
    let curPos = 0, curNeg = 0, curTotal = 0;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === todayStr) {
        summaryRow = i + 1;
        curPos = parseFloat(rows[i][3]) || 0;
        curNeg = parseFloat(rows[i][4]) || 0;
        curTotal = parseFloat(rows[i][5]) || 0;
        break;
      }
    }

    const isNewRow = summaryRow === -1;
    if (isNewRow) {
      summaryRow = state.weekly.lastSummaryRow + 1;
      state.weekly.lastSummaryRow = summaryRow;
    }

    const newPos = positiveScore > 0 ? curPos + positiveScore : curPos;
    const newNeg = negativeScore < 0 ? curNeg + negativeScore : curNeg;
    const newTotal = curTotal + positiveScore + negativeScore;

    // Queue all writes. They share one sync.
    if (isNewRow) {
      summarySheet.getRange(`${CONFIG.SUMMARY.DATE_COLUMN}${summaryRow}`).values = [[todayStr]];
    }
    if (positiveScore > 0) {
      summarySheet.getRange(`${CONFIG.SUMMARY.POSITIVE_SCORE_COLUMN}${summaryRow}`).values = [[newPos]];
    }
    if (negativeScore < 0) {
      summarySheet.getRange(`${CONFIG.SUMMARY.NEGATIVE_SCORE_COLUMN}${summaryRow}`).values = [[newNeg]];
    }
    summarySheet.getRange(`${CONFIG.SUMMARY.TOTAL_SCORE_COLUMN}${summaryRow}`).values = [[newTotal]];
    await context.sync();
  } catch (error) {
    // Sheet absent or any unexpected read/write failure: log and
    // return silently. This matches the pre-task-#37 contract of
    // "no-op when Summary sheet is missing".
    if (!/not found/i.test(String(error && error.message))) {
      console.error('Update summary error:', error);
    }
  }
}

// ============================================================================
// ARCHIVE & NEW WEEK FUNCTIONS
// ============================================================================

/**
 * Export Summary sheet data as CSV
 */
async function exportSummaryData() {
  await withStatus('Export summary', async () => {
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
  });
}

// Export for use in other modules
window.initializeSummarySheet = initializeSummarySheet;
window.updateSummary = updateSummary;
window.exportSummaryData = exportSummaryData;

/**
 * Read today's running scores from the Summary sheet.
 * Returns { positive, negative, total } or null if today's row
 * isn't recorded yet. Does NOT throw on missing Summary sheet —
 * returns null instead.
 *
 * Uses the same row-lookup as updateSummary, but read-only and
 * cheap (one sync to find the row, one to read the three cells).
 *
 * @param {Excel.RequestContext} context - Excel context owned by caller
 * @returns {Promise<{positive: number, negative: number, total: number} | null>}
 */
async function getTodayScore(context) {
  const summarySheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.SUMMARY_SHEET);
  await context.sync();
  if (summarySheet.isNullObject) return null;

  const todayStr = formatDateYYYYMMDD(new Date());
  const summaryRange = summarySheet.getRange(
    `${CONFIG.SUMMARY.DATE_COLUMN}1:${CONFIG.SUMMARY.DATE_COLUMN}${state.weekly.lastSummaryRow + 1}`
  );
  summaryRange.load('values');
  await context.sync();

  let summaryRow = -1;
  for (let i = 0; i < summaryRange.values.length; i++) {
    if (String(summaryRange.values[i][0]) === todayStr) {
      summaryRow = i + 1;
      break;
    }
  }
  if (summaryRow === -1) return null;

  const posCell = summarySheet.getRange(`${CONFIG.SUMMARY.POSITIVE_SCORE_COLUMN}${summaryRow}`);
  const negCell = summarySheet.getRange(`${CONFIG.SUMMARY.NEGATIVE_SCORE_COLUMN}${summaryRow}`);
  const totalCell = summarySheet.getRange(`${CONFIG.SUMMARY.TOTAL_SCORE_COLUMN}${summaryRow}`);
  posCell.load('values');
  negCell.load('values');
  totalCell.load('values');
  await context.sync();

  const toNumber = (v) => {
    const n = parseFloat(v);
    return Number.isFinite(n) ? n : 0;
  };
  return {
    positive: toNumber(posCell.values[0][0]),
    negative: toNumber(negCell.values[0][0]),
    total: toNumber(totalCell.values[0][0]),
  };
}

window.getTodayScore = getTodayScore;
