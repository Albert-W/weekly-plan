/**
 * Utility functions for the Combined Tracker Add-in
 *
 * This file contains helper functions for date formatting,
 * column conversion, and other common operations.
 */

/**
 * Format date as YYYYMMDD string
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDateYYYYMMDD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + m + d;
}

/**
 * Format date and time as "YYYYMMDD HH:MM:SS" string
 * @param {Date} date - The date to format
 * @returns {string} Formatted datetime string
 */
function formatDateTime(date) {
  return formatDateYYYYMMDD(date) + ' ' +
    String(date.getHours()).padStart(2, '0') + ':' +
    String(date.getMinutes()).padStart(2, '0') + ':' +
    String(date.getSeconds()).padStart(2, '0');
}

/**
 * Convert column letter(s) to 0-based index
 * @param {string} letter - Column letter (e.g., 'A', 'AA')
 * @returns {number} 0-based column index
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0));
  }
  return index;
}

/**
 * Convert 0-based index to column letter(s)
 * @param {number} index - 0-based column index
 * @returns {string} Column letter (e.g., 'A', 'AA')
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
 * Calculate the Monday of the current week
 * @param {Date} date - Reference date
 * @returns {Date} The Monday of the week containing the reference date
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
 * Calculate days between two dates
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @returns {number} Number of days between dates
 */
function daysBetween(date1, date2) {
  const oneDay = 24 * 60 * 60 * 1000;
  return Math.floor((date2 - date1) / oneDay);
}

/**
 * Parse Excel cell address into column and row
 * @param {string} address - Cell address (e.g., 'A1', 'AB123')
 * @returns {Object|null} Object with column, colIndex, and row, or null if invalid
 */
function parseAddress(address) {
  const match = address.match(/([A-Z]+)(\d+)/);
  if (!match) return null;

  return {
    column: match[1],
    colIndex: columnLetterToIndex(match[1]) + 1, // 1-based
    row: parseInt(match[2])
  };
}

// Export for use in other modules
window.formatDateYYYYMMDD = formatDateYYYYMMDD;
window.formatDateTime = formatDateTime;
window.columnLetterToIndex = columnLetterToIndex;
window.indexToColumnLetter = indexToColumnLetter;
window.getMonday = getMonday;
window.daysBetween = daysBetween;
window.parseAddress = parseAddress;
