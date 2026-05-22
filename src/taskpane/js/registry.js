/**
 * Sheet-handler registry.
 *
 * Each domain module (habits.js, weekly.js, ...) can register
 * optional per-sheet callbacks here. events.js dispatches by
 * sheet name without needing to know which sheets exist.
 *
 * Handler shape:
 *   {
 *     onSelection: async (context, address, column, colIndex, row) => {},
 *     onActivate:  async (context, sheetName) => {},
 *     onChange:    async (context, address, colIndex, row) => {},
 *   }
 *
 * All callbacks are optional; missing ones simply aren't called.
 *
 * Usage from a domain module:
 *   registerSheetHandlers(CONFIG.WEEKLY_SHEET, {
 *     onSelection: handleWeeklySelection,
 *     onActivate:  initializeWeeklyOnOpen,
 *     onChange:    handleWeeklyCellChange,
 *   });
 */

const sheetHandlers = {};

function registerSheetHandlers(sheetName, handlers) {
  sheetHandlers[sheetName] = { ...(sheetHandlers[sheetName] || {}), ...handlers };
}

function getSheetHandlers(sheetName) {
  return sheetHandlers[sheetName] || null;
}

// Export for use in other modules
window.sheetHandlers = sheetHandlers;
window.registerSheetHandlers = registerSheetHandlers;
window.getSheetHandlers = getSheetHandlers;
