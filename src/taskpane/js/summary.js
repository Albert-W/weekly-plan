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
    const summaryRange = summarySheet.getRange(`${CONFIG.SUMMARY.DATE_COLUMN}1:${CONFIG.SUMMARY.DATE_COLUMN}${state.weekly.lastSummaryRow + 1}`);
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
      summaryRow = state.weekly.lastSummaryRow + 1;
      summarySheet.getRange(`${CONFIG.SUMMARY.DATE_COLUMN}${summaryRow}`).values = [[todayStr]];
      state.weekly.lastSummaryRow = summaryRow;
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
window.initializeSummarySheet = initializeSummarySheet;
window.updateSummary = updateSummary;
window.exportSummaryData = exportSummaryData;
