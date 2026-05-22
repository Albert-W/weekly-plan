/**
 * Test harness — pulls the globals attached by tests/setup.js (which
 * loaded the src/ files into globalThis) and re-exports them as named
 * imports, so test files written in strict-mode ESM can refer to them
 * normally instead of using `window.xxx` everywhere.
 */

// Configuration & state
export const CONFIG = globalThis.CONFIG;
export const state = globalThis.state;

// utils.js
export const formatDateYYYYMMDD = globalThis.formatDateYYYYMMDD;
export const formatDateTime = globalThis.formatDateTime;
export const columnLetterToIndex = globalThis.columnLetterToIndex;
export const indexToColumnLetter = globalThis.indexToColumnLetter;
export const getMonday = globalThis.getMonday;
export const daysBetween = globalThis.daysBetween;
export const parseAddress = globalThis.parseAddress;

// weekly.js
export const buildWeeklyCSV = globalThis.buildWeeklyCSV;
export const formatExcelTime = globalThis.formatExcelTime;
export const escapeCSV = globalThis.escapeCSV;
export const processWeeklyScoreChange = globalThis.processWeeklyScoreChange;
export const clearForNewWeek = globalThis.clearForNewWeek;
export const updateSummary = globalThis.updateSummary;

// habits.js
export const recordHabitDone = globalThis.recordHabitDone;

// events.js
export const registerOnChangedEvent = globalThis.registerOnChangedEvent;
export const handleCellChanged = globalThis.handleCellChanged;
