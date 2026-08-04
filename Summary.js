/**
 * Summary-sheet domain logic (Google Sheets edition).
 *
 * Ported from src/taskpane/js/summary.js. Accumulates today's positive
 * and negative score deltas into D/E/F of today's row (creating it if
 * needed). The Office build guarded the read-modify-write with a
 * per-sheet promise chain; here we use LockService so rapid edits can't
 * lose increments.
 */

/**
 * Update the Summary sheet with today's score deltas.
 * Silent no-op when the Summary sheet is absent.
 * @param {number} positiveScore accumulated into col D when > 0
 * @param {number} negativeScore accumulated into col E when < 0
 */
function updateSummary(positiveScore, negativeScore) {
  const sheet = getSheetByName_(CONFIG.SUMMARY_SHEET);
  if (!sheet) return;

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('updateSummary: could not obtain lock: ' + e.message);
    return;
  }

  try {
    const S = CONFIG.SUMMARY;
    const lastRow = Math.max(getLastRowInColumn_(sheet, 1), 1);
    const rows = sheet.getRange('A1:F' + lastRow).getValues();

    const todayStr = formatDateYYYYMMDD(new Date());
    let summaryRow = -1;
    let curPos = 0;
    let curNeg = 0;
    let curTotal = 0;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === todayStr) {
        summaryRow = i + 1; // 1-based
        curPos = parseFloat(rows[i][3]) || 0;
        curNeg = parseFloat(rows[i][4]) || 0;
        curTotal = parseFloat(rows[i][5]) || 0;
        break;
      }
    }

    const isNewRow = summaryRow === -1;
    if (isNewRow) summaryRow = lastRow + 1;

    const newPos = positiveScore > 0 ? curPos + positiveScore : curPos;
    const newNeg = negativeScore < 0 ? curNeg + negativeScore : curNeg;
    const newTotal = curTotal + positiveScore + negativeScore;

    if (isNewRow) {
      sheet.getRange(S.DATE_COLUMN + summaryRow).setValue(todayStr);
    }
    if (positiveScore > 0) {
      sheet.getRange(S.POSITIVE_SCORE_COLUMN + summaryRow).setValue(newPos);
    }
    if (negativeScore < 0) {
      sheet.getRange(S.NEGATIVE_SCORE_COLUMN + summaryRow).setValue(newNeg);
    }
    sheet.getRange(S.TOTAL_SCORE_COLUMN + summaryRow).setValue(newTotal);
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log('Update summary error: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Read the running scores for a specific date from the Summary sheet.
 * @param {string} dateStr YYYYMMDD
 * @returns {{positive: number, negative: number, total: number}|null}
 *   null when the sheet is missing or that date's row isn't recorded.
 */
function getSummaryForDate_(dateStr) {
  const sheet = getSheetByName_(CONFIG.SUMMARY_SHEET);
  if (!sheet) return null;

  const lastRow = Math.max(getLastRowInColumn_(sheet, 1), 1);
  const rows = sheet.getRange('A1:F' + lastRow).getValues();

  const toNumber = (v) => {
    const n = parseFloat(v);
    return isFinite(n) ? n : 0;
  };

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === dateStr) {
      return {
        positive: toNumber(rows[i][3]),
        negative: toNumber(rows[i][4]),
        total: toNumber(rows[i][5]),
      };
    }
  }
  return null;
}

/**
 * Read today's running scores from the Summary sheet.
 * @returns {{positive: number, negative: number, total: number}|null}
 *   null when the sheet is missing or today's row isn't recorded yet.
 */
function getTodayScore() {
  return getSummaryForDate_(formatDateYYYYMMDD(new Date()));
}

/**
 * Export the Summary sheet as a CSV file in the Drive archive folder.
 * @returns {string|null} the Drive file URL, or null if nothing exported
 */
function exportSummaryData() {
  const sheet = getSheetByName_(CONFIG.SUMMARY_SHEET);
  if (!sheet) {
    toast_('Summary sheet not found.', 'Weekly Plan');
    return null;
  }
  const csv = sheetToCsv_(sheet);
  if (!csv) return null;
  const filename = 'Summary_' + formatDateYYYYMMDD(new Date()) + '.csv';
  const url = saveCsvToDrive_(csv, filename);
  toast_('Summary exported to Drive.', 'Weekly Plan');
  return url;
}
