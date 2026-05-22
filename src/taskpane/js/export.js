/**
 * Export / archive helpers for the Combined Tracker Add-in.
 *
 * Encapsulates CSV serialization (buildWeeklyCSV, exportSheetAsCSV,
 * escapeCSV, formatExcelTime), file downloads (downloadCSV), and the
 * user-facing archive/export entry points (exportWeekData,
 * archiveWeekAutomatically, archiveAndStartNewWeek, exportWeeklyAsXLS).
 *
 * Depends on config, state, utils. Loaded BEFORE weekly.js because
 * weekly.js (and habits.js indirectly via updateSummary) call into
 * these helpers.
 */

/**
 * Format an Excel time-cell value as "HH:MM".
 * Excel stores times either as a fraction of a day (0.645833 = 15:30)
 * or as a string. Anything else is returned as-is via String().
 *
 * @param {number|string|*} time - Cell value
 * @returns {string}
 */
function formatExcelTime(time) {
  if (typeof time === 'number') {
    const hours = Math.floor(time * 24);
    const mins = Math.round((time * 24 - hours) * 60);
    return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
  }
  return String(time);
}

/**
 * Escape a value for CSV (handle commas, quotes, newlines)
 * @param {string} value - Value to escape
 * @returns {string} Escaped value
 */

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

/**
 * Download a string as a CSV file.
 *
 * Browser path: Excel Online runs the task pane in a normal browser
 * iframe, so the standard <a download> trick works and saves the file
 * to the user's Downloads folder.
 *
 * Excel Desktop path: the task pane is a sandboxed webview that
 * blocks the implicit save dialog, so the click() silently no-ops.
 * Instead we copy the CSV to the clipboard (when available) and
 * surface a clear status message so the user knows what happened.
 *
 * @param {string} content - CSV content
 * @param {string} filename - Filename for the download
 * @returns {boolean} true if a real download was attempted; false on
 *   the Desktop fallback path.
 */
function downloadCSV(content, filename) {
  if (!isExcelOnline()) {
    let copied = false;
    if (typeof navigator !== 'undefined' &&
        navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(content).then(
        () => { /* ok */ },
        (err) => { console.warn('Clipboard write failed:', err); }
      );
      copied = true;
    }
    const msg = copied
      ? `📋 Excel Desktop can't auto-download. CSV copied to clipboard — paste into a new file and save as "${filename}".`
      : `⚠️ Excel Desktop can't auto-download "${filename}". Use Excel Online for one-click CSV export.`;
    showStatus(msg, 'warning');
    console.log('Desktop CSV fallback for:', filename);
    return false;
  }

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
  return true;
}

/**
 * Show instructions for creating a copy in OneDrive
 */

/**
 * Build a CSV snapshot of the current Weekly sheet.
 *
 * Single source of truth for both the auto-archive on new-week
 * detection and the manual "Export All" action. Caller owns the
 * Excel.run context so this composes cleanly with other operations.
 *
 * @param {Excel.RequestContext} context - Excel context
 * @returns {Promise<{csv: string, filename: string}>}
 */
async function buildWeeklyCSV(context) {
  const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);

  const dateCell = weeklySheet.getRange('B4');
  const headerRange = weeklySheet.getRange('D4:P4');
  const timeRange = weeklySheet.getRange(
    `B${CONFIG.WEEKLY.DATA_START_ROW}:B${CONFIG.WEEKLY.LAST_TIME_ROW}`
  );
  const dataRange = weeklySheet.getRange(
    `C${CONFIG.WEEKLY.DATA_START_ROW}:P${CONFIG.WEEKLY.LAST_TIME_ROW}`
  );

  dateCell.load('values');
  headerRange.load('values');
  timeRange.load('values');
  dataRange.load('values');
  await context.sync();

  // Build week label for filename
  const dateStr = String(dateCell.values[0][0] || '');
  const headerValues = headerRange.values[0];
  const firstDay = headerValues[0];
  const lastDay = headerValues[headerValues.length - 1];
  const weekLabel = `${dateStr.replace(' ', '-')}_${firstDay}-${lastDay}`;
  const filename = `Weekly_${weekLabel}.csv`;

  // Build CSV header
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const headers = ['Time'];
  for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
    headers.push(`${days[d]}_Task`);
    headers.push(`${days[d]}_Score`);
  }
  const lines = [headers.join(',')];

  // Build CSV rows
  for (let i = 0; i < timeRange.values.length; i++) {
    const time = timeRange.values[i][0];
    if (time === '' || time === null) continue;

    const row = [escapeCSV(formatExcelTime(time))];

    for (let d = 0; d < CONFIG.WEEKLY.DAYS_IN_WEEK; d++) {
      const taskCol = d * 2;
      const scoreCol = d * 2 + 1;
      const task = dataRange.values[i][taskCol] || '';
      const score = dataRange.values[i][scoreCol];
      row.push(escapeCSV(String(task)));
      row.push(score !== null && score !== '' ? score : '');
    }
    lines.push(row.join(','));
  }

  // Trailing newline matches the original output exactly.
  const csv = lines.join('\n') + '\n';
  return { csv, filename };
}

/**
 * Format an Excel time-cell value as "HH:MM".
 * Excel stores times either as a fraction of a day (0.645833 = 15:30)
 * or as a string. Anything else is returned as-is via String().
 *
 * @param {number|string|*} time - Cell value
 * @returns {string}
 */

/**
 * Automatically archive week data when new week is detected.
 * Called from initializeWeeklyOnOpen.
 */
async function archiveWeekAutomatically() {
  try {
    let result = null;
    await Excel.run(async (context) => {
      result = await buildWeeklyCSV(context);
    });

    if (result && result.csv.split('\n').length > 2) {
      downloadCSV(result.csv, result.filename);
      console.log('✅ Week archived to:', result.filename);
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

/**
 * Export current week data to CSV format.
 * Thin wrapper over buildWeeklyCSV that owns its own Excel.run.
 * @returns {Promise<{csv: string, filename: string} | null>}
 */
async function exportWeekData() {
  try {
    let result = null;
    await Excel.run(async (context) => {
      result = await buildWeeklyCSV(context);
    });
    return result;
  } catch (error) {
    console.error('Export error:', error);
    showStatus('Error exporting: ' + error.message, 'error');
    return null;
  }
}

/**
 * Build a CSV snapshot of the current Weekly sheet.
 *
 * Single source of truth for both the auto-archive on new-week
 * detection and the manual "Export All" action. Caller owns the
 * Excel.run context so this composes cleanly with other operations.
 *
 * @param {Excel.RequestContext} context - Excel context
 * @returns {Promise<{csv: string, filename: string}>}
 */

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

/**
 * Archive the current week's data and start a new week
 * This exports data as CSV, then clears for new week
 */
async function archiveAndStartNewWeek() {
  await withStatus('Archive week', async () => {
    showStatus('📦 Archiving week data...', 'info');
    const weekData = await exportWeekData();
    if (weekData) {
      downloadCSV(weekData.csv, weekData.filename);
      showArchiveInstructions();
    }
    showStatus('📥 Week archived! Click "Start New Week" to clear data.', 'success');
  });
}

/**
 * Export current week data to CSV format.
 * Thin wrapper over buildWeeklyCSV that owns its own Excel.run.
 * @returns {Promise<{csv: string, filename: string} | null>}
 */

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

/**
 * Export all important sheets as separate CSV files for backup
 * This exports Weekly, Goals, and Charter sheets
 */
async function exportWeeklyAsXLS() {
  await withStatus('Export all sheets', async () => {
    showStatus('📊 Exporting all sheets...', 'info');

    let weekLabel = '';

    await Excel.run(async (context) => {
      const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      const dateCell = weeklySheet.getRange('B4');
      dateCell.load('values');
      const headerRange = weeklySheet.getRange('D4:P4');
      headerRange.load('values');
      await context.sync();

      const dateStr = String(dateCell.values[0][0] || '').trim();
      if (dateStr && headerRange.values[0].length > 0) {
        const firstDay = headerRange.values[0][0];
        const lastDay = headerRange.values[0][headerRange.values[0].length - 1];
        weekLabel = `${dateStr.replace(' ', '-')}_${firstDay}-${lastDay}`;
      } else {
        weekLabel = formatDateYYYYMMDD(new Date());
      }
    });

    const sheetsToExport = [
      { name: CONFIG.WEEKLY_SHEET, prefix: 'Weekly' },
      { name: CONFIG.GOALS_SHEET, prefix: 'Goals' },
      { name: CONFIG.CHARTER_SHEET, prefix: 'Charter' }
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
  });
}

/**
 * Export a single sheet as CSV
 * @param {string} sheetName - Name of the sheet to export
 * @returns {string|null} CSV content or null if sheet not found
 */

// Export for use in other modules
window.formatExcelTime = formatExcelTime;
window.escapeCSV = escapeCSV;
window.downloadCSV = downloadCSV;
window.buildWeeklyCSV = buildWeeklyCSV;
window.archiveWeekAutomatically = archiveWeekAutomatically;
window.exportWeekData = exportWeekData;
window.showArchiveInstructions = showArchiveInstructions;
window.archiveAndStartNewWeek = archiveAndStartNewWeek;
window.exportSheetAsCSV = exportSheetAsCSV;
window.exportWeeklyAsXLS = exportWeeklyAsXLS;
