/**
 * Export / archive helpers (Google Sheets edition).
 *
 * Ported from src/taskpane/js/export.js. Keeps the OWASP CSV
 * formula-injection guard. The Office build downloaded CSVs via an
 * <a download> link / clipboard; GAS can't trigger a browser download
 * from the server, so CSVs are written to a Drive folder, and finished
 * weeks are additionally appended to an in-spreadsheet "Archive" tab.
 */

/**
 * Format a time cell value as "HH:MM".
 * Sheets returns time-of-day cells as Date objects (or fractions of a
 * day for some imports); our scaffold stores plain "HH:MM" strings.
 * @param {Date|number|string|*} time
 * @returns {string}
 */
function formatTimeCell_(time) {
  if (time instanceof Date) {
    return (
      String(time.getHours()).padStart(2, '0') +
      ':' +
      String(time.getMinutes()).padStart(2, '0')
    );
  }
  if (typeof time === 'number' && time > 0 && time < 1) {
    const hours = Math.floor(time * 24);
    const mins = Math.round((time * 24 - hours) * 60);
    return String(hours).padStart(2, '0') + ':' + String(mins).padStart(2, '0');
  }
  return String(time);
}

/**
 * Escape a value for CSV, with OWASP formula-injection defense for
 * string cells starting with = + - @ tab or CR.
 * @param {*} value
 * @returns {string}
 */
function escapeCSV(value) {
  if (value === null || value === undefined) return '';
  let str = String(value);

  if (typeof value === 'string' && /^[=+\-@\t\r]/.test(str)) {
    str = "'" + str;
  }

  if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * Serialize a whole sheet's used range to CSV (time cells formatted).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {string}
 */
function sheetToCsv_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const lines = [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const cells = [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell instanceof Date) {
        cells.push(escapeCSV(formatTimeCell_(cell)));
      } else {
        cells.push(escapeCSV(cell));
      }
    }
    lines.push(cells.join(','));
  }
  return lines.join('\n') + '\n';
}

/**
 * Build a CSV snapshot of the Weekly sheet plus structured rows for the
 * Archive tab.
 * @returns {{csv: string, filename: string, weekLabel: string, dataRows: Array<Array>}}
 */
function buildWeeklyData_() {
  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) throw new Error('Weekly sheet not found');
  const W = CONFIG.WEEKLY;

  const dateStr = String(sheet.getRange(W.DATE_CELL).getValue() || '');
  const headerValues = sheet.getRange(W.HEADER_RANGE).getValues()[0];
  const firstDay = headerValues[0];
  const lastDay = headerValues[headerValues.length - 1];
  const weekLabel = dateStr.replace(' ', '-') + '_' + firstDay + '-' + lastDay;
  const filename = 'Weekly_' + weekLabel + '.csv';

  const timeValues = sheet
    .getRange(W.DATA_START_ROW, W.TIME_COLUMN, W.LAST_TIME_ROW - W.DATA_START_ROW + 1, 1)
    .getValues();
  // C..P data (task/score pairs for all 7 days).
  const dataValues = sheet
    .getRange(W.DATA_START_ROW, 3, W.LAST_TIME_ROW - W.DATA_START_ROW + 1, 14)
    .getValues();

  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const headers = ['Time'];
  for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
    headers.push(days[d] + '_Task');
    headers.push(days[d] + '_Score');
  }

  const lines = [headers.join(',')];
  const dataRows = []; // [time, task, score, task, score, ...] for archive
  for (let i = 0; i < timeValues.length; i++) {
    const time = timeValues[i][0];
    if (time === '' || time === null) continue;

    const timeStr = formatTimeCell_(time);
    const csvRow = [escapeCSV(timeStr)];
    const archiveRow = [timeStr];
    for (let d = 0; d < W.DAYS_IN_WEEK; d++) {
      const taskCol = d * 2;
      const scoreCol = d * 2 + 1;
      const task = dataValues[i][taskCol] || '';
      const score = dataValues[i][scoreCol];
      const scoreOut = score !== null && score !== '' ? score : '';
      csvRow.push(escapeCSV(String(task)));
      csvRow.push(scoreOut);
      archiveRow.push(task);
      archiveRow.push(scoreOut);
    }
    lines.push(csvRow.join(','));
    dataRows.push(archiveRow);
  }

  return { csv: lines.join('\n') + '\n', filename: filename, weekLabel: weekLabel, dataRows: dataRows };
}

/**
 * Append a finished week's rows to the in-spreadsheet Archive tab,
 * each prefixed with the week label.
 * @param {string} weekLabel
 * @param {Array<Array>} dataRows
 * @returns {number} number of rows appended
 */
function appendWeekToArchiveSheet_(weekLabel, dataRows) {
  if (!dataRows || dataRows.length === 0) return 0;
  const sheet = getOrCreateSheet_(CONFIG.ARCHIVE_SHEET);
  setUpArchiveSheet_(); // ensure header exists

  const startRow = sheet.getLastRow() + 1;
  const out = dataRows.map((r) => [weekLabel].concat(r));
  sheet.getRange(startRow, 1, out.length, out[0].length).setValues(out);
  return out.length;
}

/**
 * Archive the current Weekly sheet: append to Archive tab AND save CSV
 * to Drive. Called by new-week rollover.
 * @returns {{url: string|null, rows: number}}
 */
function archiveWeek_() {
  const data = buildWeeklyData_();
  const rows = appendWeekToArchiveSheet_(data.weekLabel, data.dataRows);
  let url = null;
  if (data.dataRows.length > 0) {
    url = saveCsvToDrive_(data.csv, data.filename);
  }
  return { url: url, rows: rows };
}

/**
 * Export the important sheets as CSV files in the Drive archive folder.
 * @returns {{count: number, folderUrl: string|null}}
 */
function exportAllSheets() {
  const names = [
    CONFIG.WEEKLY_SHEET,
    CONFIG.HABITS_SHEET,
    CONFIG.TASKS_SHEET,
    CONFIG.SUMMARY_SHEET,
  ];
  const weekLabel = formatDateYYYYMMDD(new Date());
  let count = 0;
  let folderUrl = null;
  for (let i = 0; i < names.length; i++) {
    const sheet = getSheetByName_(names[i]);
    if (!sheet) continue;
    const csv = sheetToCsv_(sheet);
    if (!csv) continue;
    const filename = names[i] + '_' + weekLabel + '.csv';
    const url = saveCsvToDrive_(csv, filename);
    if (url) {
      folderUrl = getOrCreateArchiveFolder_().getUrl();
      count++;
    }
  }
  toast_('Exported ' + count + ' sheet(s) to Drive.', 'Weekly Plan');
  return { count: count, folderUrl: folderUrl };
}

// ----------------------------------------------------------------------
// Drive helpers
// ----------------------------------------------------------------------

/**
 * Get (or create) the Drive folder where CSV archives are saved.
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateArchiveFolder_() {
  const name = CONFIG.DRIVE_ARCHIVE_FOLDER;
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/**
 * Save CSV content to a file in the archive folder.
 * @param {string} content
 * @param {string} filename
 * @returns {string} the Drive file URL
 */
function saveCsvToDrive_(content, filename) {
  const folder = getOrCreateArchiveFolder_();
  const file = folder.createFile(filename, content, MimeType.CSV);
  return file.getUrl();
}
