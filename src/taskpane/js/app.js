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

  // DOM bootstrap runs unconditionally so the date icon shows and
  // buttons are wired even if we're not in Excel.
  bootstrapDom();

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

/**
 * One-time DOM bootstrap. Sets the date icon and wires buttons by
 * `data-action` / `data-toggle` attributes instead of inline onclick.
 *
 * Adding a new button means: drop it in the HTML with the right
 * data-attribute, then add a matching entry to BUTTON_ACTIONS below.
 * No more inline JavaScript in taskpane.html.
 */
const BUTTON_ACTIONS = {
  'refresh':              () => refreshCurrentSheet(),
  'sort-habits':          () => sortHabits(),
  'refresh-habit-dates':  () => refreshHabitsDates(),
  'random-pick':          () => randomPickFromUI(),
  'export-all':           () => exportWeeklyAsXLS(),
  'start-new-week':       () => startNewWeekFromUI(),
  'export-summary':       () => exportSummaryData(),
  'add-task':             () => addTask(),
};

function bootstrapDom() {
  // Date icon — was an inline <script> in taskpane.html.
  const dateNum = document.getElementById('date-num');
  if (dateNum) dateNum.textContent = new Date().getDate();

  // Action buttons.
  for (const btn of document.querySelectorAll('button[data-action]')) {
    const action = btn.getAttribute('data-action');
    const fn = BUTTON_ACTIONS[action];
    if (!fn) {
      console.warn('No handler for data-action="' + action + '"');
      continue;
    }
    btn.addEventListener('click', fn);
  }

  // Toggle buttons.
  for (const btn of document.querySelectorAll('button[data-toggle]')) {
    const targetId = btn.getAttribute('data-toggle');
    btn.addEventListener('click', () => toggleSection(targetId));
  }
}

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

    // Start the background ticker so the time-row highlight follows
    // the real clock instead of only updating on user actions.
    startTimeHighlightTicker();

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
 * Also re-initializes Weekly sheet if it's a new day
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

      // Check if we need to re-initialize Weekly sheet (new day check)
      if (sheetName === CONFIG.WEEKLY_SHEET) {
        const today = formatDateYYYYMMDD(new Date());
        const lastInit = state.weekly.lastInitDate;

        console.log('Weekly sheet refresh. Today:', today, 'Last init:', lastInit);

        if (lastInit !== today) {
          console.log('🌅 New day detected! Re-initializing Weekly sheet...');
          await initializeWeeklyOnOpen(context);
        } else {
          // Same day - just refresh time highlighting
          const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
          await highlightCurrentTimeRow(context, weeklySheet);
        }
      }

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
