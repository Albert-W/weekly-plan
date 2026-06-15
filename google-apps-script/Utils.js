/**
 * Utility functions for the Weekly Plan Google Sheets edition.
 *
 * Ported from src/taskpane/js/utils.js. The pure date/column helpers
 * are copied verbatim (they are framework-agnostic JS). Office-only
 * helpers (parseAddress, isExcelOnline) are dropped; small Apps Script
 * sheet helpers are added at the bottom.
 */

/**
 * Format date as YYYYMMDD string.
 * @param {Date} date
 * @returns {string}
 */
function formatDateYYYYMMDD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + m + d;
}

/**
 * Format date and time as "YYYYMMDD HH:MM:SS".
 * @param {Date} date
 * @returns {string}
 */
function formatDateTime(date) {
  return (
    formatDateYYYYMMDD(date) +
    ' ' +
    String(date.getHours()).padStart(2, '0') +
    ':' +
    String(date.getMinutes()).padStart(2, '0') +
    ':' +
    String(date.getSeconds()).padStart(2, '0')
  );
}

/**
 * Convert column letter(s) to a 0-based index.
 * Base-26 with no zero digit (A..Z, AA..AZ, ...).
 * @param {string} letter e.g. 'A', 'Z', 'AA', 'AZ'
 * @returns {number} 0-based index ('A' -> 0, 'AA' -> 26)
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

/**
 * Convert a 0-based index to column letter(s).
 * @param {number} index 0-based column index
 * @returns {string}
 */
function indexToColumnLetter(index) {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 'A'.charCodeAt(0)) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

/**
 * Monday of the week containing `date`.
 * @param {Date} date
 * @returns {Date}
 */
function getMonday(date) {
  const d = new Date(date);
  const dayOfWeek = d.getDay();
  const diff = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  d.setDate(d.getDate() - diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Whole days between two dates.
 * @param {Date} date1
 * @param {Date} date2
 * @returns {number}
 */
function daysBetween(date1, date2) {
  const oneDay = 24 * 60 * 60 * 1000;
  return Math.floor((date2 - date1) / oneDay);
}

/**
 * 1-based task column number for a day index.
 * Day 0 (Mon) = col 3 (C) ... Day 6 (Sun) = col 15 (O).
 * @param {number} dayIndex 0 = Monday ... 6 = Sunday
 * @returns {number}
 */
function getTaskColForDay(dayIndex) {
  return dayIndex * 2 + 3;
}

/**
 * 1-based score column number for a day index.
 * Day 0 (Mon) = col 4 (D) ... Day 6 (Sun) = col 16 (P).
 * @param {number} dayIndex 0 = Monday ... 6 = Sunday
 * @returns {number}
 */
function getScoreColForDay(dayIndex) {
  return dayIndex * 2 + 4;
}

/**
 * Task column letter for a day index (0 -> 'C').
 * @param {number} dayIndex
 * @returns {string}
 */
function getTaskColLetterForDay(dayIndex) {
  return indexToColumnLetter(getTaskColForDay(dayIndex) - 1);
}

/**
 * Score column letter for a day index (0 -> 'D').
 * @param {number} dayIndex
 * @returns {string}
 */
function getScoreColLetterForDay(dayIndex) {
  return indexToColumnLetter(getScoreColForDay(dayIndex) - 1);
}

/**
 * Current day index where 0 = Monday ... 6 = Sunday.
 * Replaces the persisted state.weekly.currentDayIndex (GAS is stateless).
 * @param {Date} [date]
 * @returns {number}
 */
function getCurrentDayIndex(date) {
  const d = date || new Date();
  const dow = d.getDay(); // 0 = Sun
  return dow === 0 ? 6 : dow - 1;
}

// ----------------------------------------------------------------------
// Apps Script sheet helpers (trailing underscore = private convention)
// ----------------------------------------------------------------------

/**
 * Active spreadsheet shortcut.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Get a sheet by name, or null if it does not exist.
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheetByName_(name) {
  return getSpreadsheet_().getSheetByName(name);
}

/**
 * Get a sheet by name, creating it if missing.
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Last non-empty row in a single column (1-based). Returns 0 when empty.
 * Replaces the various used-range row-count reads from the Office build.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} col 1-based column index
 * @returns {number}
 */
function getLastRowInColumn_(sheet, col) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, col, maxRows, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    if (v !== '' && v !== null) return i + 1;
  }
  return 0;
}

/**
 * Show a transient status message to the user (toast). The GAS
 * equivalent of the Office build's showStatus banner.
 * @param {string} message
 * @param {string} [title]
 */
function toast_(message, title) {
  try {
    getSpreadsheet_().toast(message, title || 'Weekly Plan', 5);
  } catch (e) {
    // toast can fail in non-UI contexts (e.g. time-driven trigger). Log only.
    Logger.log('[toast] ' + (title || '') + ': ' + message);
  }
}

/**
 * Run `fn`, swallowing and logging any error. Returns fn() or null.
 * GAS equivalent of ui.js safeInit — for optional init steps that must
 * not abort the caller.
 * @param {string} label
 * @param {Function} fn
 * @returns {*}
 */
function safeInit_(label, fn) {
  try {
    return fn();
  } catch (e) {
    Logger.log(label + ': ' + (e && e.message ? e.message : e));
    return null;
  }
}
