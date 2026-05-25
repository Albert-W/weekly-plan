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
export const getTaskColForDay = globalThis.getTaskColForDay;
export const getScoreColForDay = globalThis.getScoreColForDay;
export const getTaskColLetterForDay = globalThis.getTaskColLetterForDay;
export const getScoreColLetterForDay = globalThis.getScoreColLetterForDay;

// weekly.js
export const buildWeeklyCSV = globalThis.buildWeeklyCSV;
export const formatExcelTime = globalThis.formatExcelTime;
export const escapeCSV = globalThis.escapeCSV;
export const processWeeklyScoreChange = globalThis.processWeeklyScoreChange;
export const clearForNewWeek = globalThis.clearForNewWeek;
// summary.js
export const updateSummary = globalThis.updateSummary;
export const getTodayScore = globalThis.getTodayScore;
export const handleWeeklySelection = globalThis.handleWeeklySelection;
export const exportSheetAsCSV = globalThis.exportSheetAsCSV;

// habits.js
export const recordHabitDone = globalThis.recordHabitDone;

// tasks.js
export const createTask = globalThis.createTask;

// events.js
export const registerOnChangedEvent = globalThis.registerOnChangedEvent;
export const handleCellChanged = globalThis.handleCellChanged;

// More from weekly.js exposed for tests
export const initializeWeeklyOnOpen = globalThis.initializeWeeklyOnOpen;
export const highlightCurrentDay = globalThis.highlightCurrentDay;
export const refreshTimeHighlight = globalThis.refreshTimeHighlight;

// habits.js
export const initializeHabitsSheet = globalThis.initializeHabitsSheet;

// ui.js
export const addTask = globalThis.addTask;
export const showModal = globalThis.showModal;
export const showWarningPopup = globalThis.showWarningPopup;
export const showInfoPopup = globalThis.showInfoPopup;

// concurrency.js
export const serializeSheetWrite = globalThis.serializeSheetWrite;
export const resetWriteChains = globalThis.resetWriteChains;
export const getInFlightChain = globalThis.getInFlightChain;

// registry.js
export const registerSheetHandlers = globalThis.registerSheetHandlers;
export const getSheetHandlers = globalThis.getSheetHandlers;

// ui.js (helpers)
export const withStatus = globalThis.withStatus;
export const safeInit = globalThis.safeInit;
