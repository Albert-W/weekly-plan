/**
 * Event handlers for the Combined Tracker Add-in
 *
 * This file contains all Office.js event handlers for
 * sheet activation, selection changes, and cell changes.
 */

/**
 * Handle sheet activation (when user clicks on a different sheet tab)
 * Also re-initializes Weekly sheet if it's a new day
 * @param {Object} event - The activation event
 */
async function handleSheetActivated(event) {
  try {
    console.log('Sheet activated event:', event);

    let newSheetName = null;

    await Excel.run(async (context) => {
      // Get the newly activated sheet
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await context.sync();

      newSheetName = activeSheet.name;
      console.log('Sheet changed to:', newSheetName);

      // Check if we need to re-initialize Weekly sheet (new day check)
      if (newSheetName === CONFIG.WEEKLY_SHEET) {
        const today = formatDateYYYYMMDD(new Date());
        const lastInit = state.weekly.lastInitDate;

        console.log('Weekly sheet activated. Today:', today, 'Last init:', lastInit);

        if (lastInit !== today) {
          console.log('🌅 New day detected! Re-initializing Weekly sheet...');
          await initializeWeeklyOnOpen(context);
        } else {
          // Same day - just refresh time highlighting
          const weeklySheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
          await highlightCurrentTimeRow(context, weeklySheet);
        }
      }

      // Only update event handlers if it's actually a different sheet
      if (state.currentSheet !== newSheetName) {
        state.currentSheet = newSheetName;

        // Re-register selection changed event for the new sheet
        await registerSelectionChangedEvent(context, activeSheet);

        // Re-register cell changed event for the new sheet
        await registerOnChangedEvent(context, activeSheet);

        await context.sync();
      }
    });

    // Update UI outside of Excel.run to ensure DOM updates
    if (newSheetName) {
      state.currentSheet = newSheetName;
    }
    updateUI();
    showStatus('Switched to: ' + state.currentSheet, 'success');

  } catch (error) {
    console.error('Sheet activation error:', error);
    showStatus('Error switching sheet: ' + error.message, 'error');
  }
}

/**
 * Register SelectionChanged event for a sheet
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - The worksheet to register for
 */
async function registerSelectionChangedEvent(context, sheet) {
  // Remove existing handler if any
  if (state.selectionHandler) {
    try {
      state.selectionHandler.remove();
      await context.sync();
    } catch (e) {
      console.log('Could not remove previous handler:', e.message);
    }
  }

  // Add new handler
  state.selectionHandler = sheet.onSelectionChanged.add(async (event) => {
    await handleSelectionChanged(event);
  });

  await context.sync();
  console.log('SelectionChanged registered for:', sheet.name);
}

/**
 * Handle selection changed - routes to appropriate handler based on sheet
 * @param {Object} event - The selection changed event
 */
async function handleSelectionChanged(event) {
  try {
    await Excel.run(async (context) => {
      const address = event.address;
      console.log('Selection:', address, 'on sheet:', state.currentSheet);

      // Parse address
      const parsed = parseAddress(address);
      if (!parsed) return;

      const { column, colIndex, row } = parsed;

      // Route to appropriate handler
      if (state.currentSheet === CONFIG.HABITS_SHEET) {
        await handleHabitsSelection(context, address, column, colIndex, row);
      } else if (state.currentSheet === CONFIG.WEEKLY_SHEET) {
        await handleWeeklySelection(context, address, column, colIndex, row);
      }
    });
  } catch (error) {
    console.error('SelectionChanged error:', error);
  }
}

/**
 * Register for cell value changes (onChanged event)
 * This is equivalent to VBA Worksheet_Change
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - The worksheet to register for
 */
async function registerOnChangedEvent(context, sheet) {
  try {
    sheet.onChanged.add(async (event) => {
      await handleCellChanged(event);
    });
    await context.sync();
    console.log('OnChanged event registered for:', sheet.name);
  } catch (e) {
    console.log('OnChanged event not supported:', e.message);
  }
}

/**
 * Handle cell value changes
 * Equivalent to VBA Worksheet_Change
 * @param {Object} event - The change event
 */
async function handleCellChanged(event) {
  try {
    await Excel.run(async (context) => {
      const address = event.address;
      console.log('Cell changed:', address, 'on sheet:', state.currentSheet);

      // Only process changes on Weekly sheet
      if (state.currentSheet !== CONFIG.WEEKLY_SHEET) return;

      // Parse address
      const parsed = parseAddress(address);
      if (!parsed) return;

      const { colIndex, row } = parsed;

      // Check if it's a score column (even: 4,6,8,10,12,14,16) in data area
      if (CONFIG.WEEKLY.SCORE_COLUMNS.includes(colIndex) &&
          row >= CONFIG.WEEKLY.DATA_START_ROW && row <= CONFIG.WEEKLY.lastTimeLine) {

        const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
        const scoreCell = sheet.getRange(address);
        scoreCell.load('values');
        await context.sync();

        const scoreValue = parseFloat(scoreCell.values[0][0]);
        if (!isNaN(scoreValue)) {
          await processWeeklyScoreChange(context, row, colIndex, scoreValue);
        }
      }
    });
  } catch (error) {
    console.error('CellChanged error:', error);
  }
}

// Export for use in other modules
window.handleSheetActivated = handleSheetActivated;
window.registerSelectionChangedEvent = registerSelectionChangedEvent;
window.handleSelectionChanged = handleSelectionChanged;
window.registerOnChangedEvent = registerOnChangedEvent;
window.handleCellChanged = handleCellChanged;
