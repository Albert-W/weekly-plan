/**
 * UI-related functions for the Combined Tracker Add-in
 *
 * This file contains functions for updating the user interface,
 * showing status messages, and toggling sections.
 */

/**
 * Update the sheet indicator in the DOM
 * @param {string} sheetName - The name of the current sheet
 */
function updateSheetIndicator(sheetName) {
  const indicator = document.getElementById('current-sheet');
  if (indicator) {
    indicator.textContent = sheetName;
    console.log('Sheet indicator updated to:', sheetName);
  } else {
    console.error('Sheet indicator element not found!');
  }
}

/**
 * Show a status message in the UI
 * @param {string} message - The message to display
 * @param {string} type - Message type: 'success', 'error', 'warning', 'info'
 */
function showStatus(message, type) {
  const el = document.getElementById('status');
  if (el) {
    el.textContent = message;
    el.className = 'status ' + (type || '');
  }
  console.log(`[${type}] ${message}`);
}

/**
 * Toggle visibility of a collapsible section
 * @param {string} id - The element ID to toggle
 */
function toggleSection(id) {
  const el = document.getElementById(id);
  if (el) {
    el.style.display = el.style.display === 'none' ? 'block' : 'none';
  }
}

/**
 * Update the UI based on current sheet
 * Shows/hides sheet-specific action sections
 */
function updateUI() {
  const isHabits = state.currentSheet === CONFIG.HABITS_SHEET;
  const isWeekly = state.currentSheet === CONFIG.WEEKLY_SHEET;

  const habitsSection = document.getElementById('habits-actions');
  const weeklySection = document.getElementById('weekly-actions');

  if (habitsSection) habitsSection.style.display = isHabits ? 'block' : 'none';
  if (weeklySection) weeklySection.style.display = isWeekly ? 'block' : 'none';

  // Update sheet indicator
  updateSheetIndicator(state.currentSheet || 'None');
}

/**
 * Add a new task to the Tasks sheet
 * Called from the Add Task form in the UI
 */
async function addTask() {
  const nameInput = document.getElementById('new-task-name');
  const weightInput = document.getElementById('new-task-weight');

  const name = nameInput ? nameInput.value.trim() : '';
  const weight = weightInput ? parseFloat(weightInput.value) || 1 : 1;

  if (!name) {
    showStatus('Please enter a task name', 'warning');
    return;
  }

  try {
    await Excel.run(async (context) => {
      const tasksSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.TASKS_SHEET);
      await context.sync();

      if (tasksSheet.isNullObject) {
        showStatus('Tasks sheet not found!', 'error');
        return;
      }

      // Find last row
      const usedRange = tasksSheet.getRange('A:A').getUsedRange();
      usedRange.load('rowCount');
      await context.sync();

      const newRow = usedRange.rowCount + 1;

      // Add task
      tasksSheet.getRange(`A${newRow}`).values = [[name]];
      tasksSheet.getRange(`B${newRow}`).values = [[weight]];
      tasksSheet.getRange(`C${newRow}`).values = [[formatDateTime(new Date())]];

      await context.sync();

      // Clear form
      if (nameInput) nameInput.value = '';
      if (weightInput) weightInput.value = '1';

      // Update state
      state.weekly.taskl = newRow;

      showStatus(`✅ Task "${name}" added!`, 'success');
    });
  } catch (error) {
    showStatus('Error: ' + error.message, 'error');
  }
}

/**
 * Prompt user to save the file
 * Note: Office.js cannot directly save files for security reasons.
 * On Excel Online, files auto-save to OneDrive/SharePoint.
 * On Desktop, we can only remind the user to save manually.
 */
function promptSaveFile() {
  // Check if we're in Excel Online or Desktop
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  if (isOnline) {
    showStatus('📁 Files auto-save on Excel Online. Your changes are saved!', 'success');
  } else {
    showStatus('💾 Please save your file: Press Ctrl+S (or Cmd+S on Mac)', 'warning');

    // Also show an alert for visibility
    if (typeof Office.context.ui !== 'undefined') {
      // Could show a dialog, but for simplicity just use status message
    }
  }
}

/**
 * Check if there are unsaved changes and prompt user
 * Call this before critical operations or on a timer
 */
function remindToSave() {
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  if (!isOnline) {
    showStatus('💡 Remember to save your file (Ctrl+S / Cmd+S)', 'info');
  }
}

/**
 * Show a warning popup dialog
 * Uses a custom modal since alert() is not supported in Office Add-ins
 * @param {string} message - The warning message to display
 */
function showWarningPopup(message) {
  // Show in status with warning style
  showStatus('⚠️ ' + message, 'warning');

  // Also show a modal dialog in the taskpane
  showModal('⚠️ Warning', message, 'warning');
}

/**
 * Show an info popup dialog
 * @param {string} title - Dialog title
 * @param {string} message - The message to display
 */
function showInfoPopup(title, message) {
  showStatus(message, 'info');
  showModal(title, message, 'info');
}

/**
 * Show a custom modal dialog in the taskpane
 * @param {string} title - Modal title
 * @param {string} message - Modal message
 * @param {string} type - 'warning', 'error', 'success', 'info'
 */
function showModal(title, message, type) {
  // Remove existing modal if any
  const existingModal = document.getElementById('custom-modal');
  if (existingModal) {
    existingModal.remove();
  }

  // Create modal HTML
  const modal = document.createElement('div');
  modal.id = 'custom-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0,0,0,0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 9999;
  `;

  const colors = {
    warning: '#e65100',
    error: '#c62828',
    success: '#2e7d32',
    info: '#1565c0'
  };

  const bgColors = {
    warning: '#fff3e0',
    error: '#ffebee',
    success: '#e8f5e9',
    info: '#e3f2fd'
  };

  modal.innerHTML = `
    <div style="
      background: white;
      border-radius: 12px;
      padding: 20px;
      max-width: 280px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      text-align: center;
    ">
      <div style="
        font-size: 18px;
        font-weight: bold;
        color: ${colors[type] || colors.info};
        margin-bottom: 12px;
      ">${title}</div>
      <div style="
        font-size: 14px;
        color: #333;
        margin-bottom: 16px;
        line-height: 1.4;
      ">${message}</div>
      <button id="modal-ok-btn" style="
        background: ${colors[type] || colors.info};
        color: white;
        border: none;
        padding: 10px 30px;
        border-radius: 8px;
        font-size: 14px;
        font-weight: bold;
        cursor: pointer;
      ">OK</button>
    </div>
  `;

  document.body.appendChild(modal);

  // Close on button click
  document.getElementById('modal-ok-btn').addEventListener('click', () => {
    modal.remove();
  });

  // Close on backdrop click
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  // Auto-close after 3 seconds
  setTimeout(() => {
    if (document.getElementById('custom-modal')) {
      modal.remove();
    }
  }, 3000);
}

// Export for use in other modules
window.updateSheetIndicator = updateSheetIndicator;
window.showStatus = showStatus;
window.toggleSection = toggleSection;
window.updateUI = updateUI;
window.addTask = addTask;
window.promptSaveFile = promptSaveFile;
window.remindToSave = remindToSave;
window.showWarningPopup = showWarningPopup;
window.showInfoPopup = showInfoPopup;
