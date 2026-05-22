/**
 * Office.js event wiring for the Combined Tracker Add-in.
 *
 * This file is pure routing. It does not know which sheets exist or
 * what their semantics are — domain modules (habits.js, weekly.js,
 * ...) self-register their per-sheet callbacks via
 * registerSheetHandlers() in registry.js. Adding a new sheet
 * requires zero changes to this file.
 */

// ----------------------------------------------------------------------
// Sheet activation
// ----------------------------------------------------------------------

/**
 * Handle sheet activation (user clicked a different sheet tab).
 * Dispatches to the activated sheet's onActivate handler (if any),
 * then re-registers selection/change listeners on the new sheet.
 */
async function handleSheetActivated(event) {
  try {
    console.log('Sheet activated event:', event);

    let newSheetName = null;

    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await context.sync();

      newSheetName = activeSheet.name;
      console.log('Sheet changed to:', newSheetName);

      // Domain-specific activation work (e.g. Weekly's new-day check).
      const handlers = getSheetHandlers(newSheetName);
      if (handlers && handlers.onActivate) {
        await handlers.onActivate(context, newSheetName);
      }

      // Only update event handlers if the sheet actually changed.
      if (state.currentSheet !== newSheetName) {
        state.currentSheet = newSheetName;
        await registerSelectionChangedEvent(context, activeSheet);
        await registerOnChangedEvent(context, activeSheet);
        await context.sync();
      }
    });

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

// ----------------------------------------------------------------------
// Selection changes
// ----------------------------------------------------------------------

/**
 * Register the SelectionChanged event for a sheet. Removes any previous
 * handler first so we never stack duplicates (see task #13 for the
 * onChanged equivalent).
 */
async function registerSelectionChangedEvent(context, sheet) {
  if (state.selectionHandler) {
    try {
      state.selectionHandler.remove();
      await context.sync();
    } catch (e) {
      console.log('Could not remove previous handler:', e.message);
    }
  }

  state.selectionHandler = sheet.onSelectionChanged.add(async (event) => {
    await handleSelectionChanged(event);
  });

  await context.sync();
  console.log('SelectionChanged registered for:', sheet.name);
}

/**
 * Dispatch a selection event to the current sheet's onSelection handler.
 */
async function handleSelectionChanged(event) {
  try {
    await Excel.run(async (context) => {
      const address = event.address;
      console.log('Selection:', address, 'on sheet:', state.currentSheet);

      const parsed = parseAddress(address);
      if (!parsed) return;
      const { column, colIndex, row } = parsed;

      const handlers = getSheetHandlers(state.currentSheet);
      if (handlers && handlers.onSelection) {
        await handlers.onSelection(context, address, column, colIndex, row);
      }
    });
  } catch (error) {
    console.error('SelectionChanged error:', error);
  }
}

// ----------------------------------------------------------------------
// Cell changes
// ----------------------------------------------------------------------

/**
 * Register the onChanged event for a sheet. Removes any previous handler
 * first to prevent stacking (regression guard from task #13).
 */
async function registerOnChangedEvent(context, sheet) {
  if (state.changeHandler) {
    try {
      state.changeHandler.remove();
      await context.sync();
    } catch (e) {
      console.log('Could not remove previous change handler:', e.message);
    }
  }

  try {
    state.changeHandler = sheet.onChanged.add(async (event) => {
      await handleCellChanged(event);
    });
    await context.sync();
    console.log('OnChanged event registered for:', sheet.name);
  } catch (e) {
    state.changeHandler = null;
    console.log('OnChanged event not supported:', e.message);
  }
}

/**
 * Dispatch a cell-changed event to the current sheet's onChange handler.
 */
async function handleCellChanged(event) {
  try {
    await Excel.run(async (context) => {
      const address = event.address;
      console.log('Cell changed:', address, 'on sheet:', state.currentSheet);

      const parsed = parseAddress(address);
      if (!parsed) return;
      const { colIndex, row } = parsed;

      const handlers = getSheetHandlers(state.currentSheet);
      if (handlers && handlers.onChange) {
        await handlers.onChange(context, address, colIndex, row);
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
