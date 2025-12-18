/**
 * Main application initialization for the Combined Tracker Add-in
 *
 * This file contains the Office.onReady handler and main
 * initialization logic that ties all modules together.
 */

// ============================================================================
// OFFICE.JS INITIALIZATION
// ============================================================================

Office.onReady((info) => {
  console.log('Office.onReady called, host:', info.host);

  if (info.host === Office.HostType.Excel) {
    console.log('Excel Add-in loaded');
    document.getElementById('sideload-msg').style.display = 'none';
    document.getElementById('app-body').style.display = 'flex';

    // Initialize the add-in
    initializeAddin();
  } else {
    console.log('Not running in Excel, host is:', info.host);
    document.getElementById('sideload-msg').innerHTML =
      '<p>This add-in only works in Excel.<br>Host detected: ' + (info.host || 'None') + '</p>';
  }
});

// ============================================================================
// MAIN INITIALIZATION
// ============================================================================

/**
 * Initialize the add-in and register event handlers
 */
async function initializeAddin() {
  console.log('initializeAddin starting...');

  // Update UI immediately to show we're trying
  updateSheetIndicator('Detecting...');

  try {
    let detectedSheetName = null;

    await Excel.run(async (context) => {
      console.log('Excel.run started');

      // Step 1: Get all sheet names first
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();

      // Get available sheet names
      const sheetNames = sheets.items.map(s => s.name);
      console.log('Available sheets:', sheetNames);

      // Step 2: Activate the Weekly sheet if it exists
      let activeSheet;
      if (sheetNames.includes(CONFIG.WEEKLY_SHEET)) {
        const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
        weeklySheet.activate();
        await context.sync();
        activeSheet = weeklySheet;
        console.log('Activated Weekly sheet on open');
      } else {
        activeSheet = context.workbook.worksheets.getActiveWorksheet();
      }

      activeSheet.load('name');
      await context.sync();

      // Store current sheet name
      detectedSheetName = activeSheet.name;
      state.currentSheet = detectedSheetName;
      console.log('Active sheet:', state.currentSheet);

      // Step 3: Initialize sheet-specific data (optional - don't fail if sheets don't exist)
      try {
        if (sheetNames.includes(CONFIG.HABITS_SHEET)) {
          await initializeHabitsSheet(context);
        }
      } catch (e) {
        console.log('Habits sheet init skipped:', e.message);
      }

      try {
        if (sheetNames.includes(CONFIG.WEEKLY_SHEET)) {
          await initializeWeeklyOnOpen(context);
        }
      } catch (e) {
        console.log('Weekly sheet init skipped:', e.message);
      }

      try {
        if (sheetNames.includes(CONFIG.TASKS_SHEET)) {
          await initializeTasksSheet(context);
        }
      } catch (e) {
        console.log('Tasks sheet init skipped:', e.message);
      }

      try {
        if (sheetNames.includes(CONFIG.SUMMARY_SHEET)) {
          await initializeSummarySheet(context);
        }
      } catch (e) {
        console.log('Summary sheet init skipped:', e.message);
      }

      // Step 4: Try to register sheet change event (may not be supported in all versions)
      try {
        context.workbook.worksheets.onActivated.add(handleSheetActivated);
        await context.sync();
        console.log('Sheet activation event registered');
      } catch (e) {
        console.log('Sheet activation event not supported:', e.message);
      }

      // Step 5: Register selection changed event
      try {
        await registerSelectionChangedEvent(context, activeSheet);
        await context.sync();
        console.log('Selection changed event registered');
      } catch (e) {
        console.log('Selection changed event failed:', e.message);
      }

      // Step 6: Register cell changed event (for score tracking)
      try {
        await registerOnChangedEvent(context, activeSheet);
        await context.sync();
        console.log('Cell changed event registered');
      } catch (e) {
        console.log('Cell changed event not supported:', e.message);
      }
    });

    console.log('Excel.run completed, sheet name:', detectedSheetName);

    // Update UI AFTER Excel.run completes - use the captured sheet name
    if (detectedSheetName) {
      state.currentSheet = detectedSheetName;
    }

    // Force update the sheet indicator
    updateSheetIndicator(state.currentSheet || 'Unknown');
    updateUI();
    showStatus('Ready! Sheet: ' + state.currentSheet, 'success');

  } catch (error) {
    console.error('Initialization error:', error);
    console.error('Error stack:', error.stack);

    // Show error but still update UI
    showStatus('Error: ' + error.message, 'error');
    updateSheetIndicator('Error');
    updateUI();
  }
}

/**
 * Manually refresh and detect current sheet
 * Call this if automatic detection fails
 */
async function refreshCurrentSheet() {
  console.log('Manual refresh triggered');
  updateSheetIndicator('Refreshing...');

  try {
    let sheetName = null;

    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await context.sync();

      sheetName = activeSheet.name;
      state.currentSheet = sheetName;
      console.log('Refreshed current sheet:', state.currentSheet);

      // Try to re-register events
      try {
        await registerSelectionChangedEvent(context, activeSheet);
        await registerOnChangedEvent(context, activeSheet);
        await context.sync();
      } catch (e) {
        console.log('Re-register events failed:', e.message);
      }
    });

    // Force update after Excel.run
    if (sheetName) {
      state.currentSheet = sheetName;
    }
    updateSheetIndicator(state.currentSheet || 'Unknown');
    updateUI();
    showStatus('Refreshed! Sheet: ' + state.currentSheet, 'success');

  } catch (error) {
    console.error('Refresh error:', error);
    updateSheetIndicator('Error');
    showStatus('Refresh failed: ' + error.message, 'error');
  }
}

/**
 * Wrapper for randomPick to be called from UI
 */
async function randomPickFromUI() {
  await Excel.run(async (context) => {
    await randomPick(context);
  });
}

// ============================================================================
// EXPOSE FUNCTIONS TO GLOBAL SCOPE
// ============================================================================

window.initializeAddin = initializeAddin;
window.refreshCurrentSheet = refreshCurrentSheet;
window.randomPickFromUI = randomPickFromUI;
