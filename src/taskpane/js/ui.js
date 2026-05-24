/**
 * UI-related functions for the Combined Tracker Add-in
 *
 * This file contains functions for updating the user interface,
 * showing status messages, and toggling sections.
 */

/**
 * Update the sheet indicator in the title bar.
 * Shows " · {sheetName}" suffixed to "Weekly Plan" in the H1, or
 * nothing if the sheet name is empty / generic placeholder.
 *
 * @param {string} sheetName - The name of the current sheet
 */
function updateSheetIndicator(sheetName) {
  const suffix = document.getElementById('current-sheet-suffix');
  if (!suffix) {
    console.error('Sheet indicator element not found!');
    return;
  }
  if (!sheetName || sheetName === 'None' || sheetName === 'Unknown') {
    suffix.textContent = '';
  } else {
    suffix.textContent = ' · ' + sheetName;
  }
  console.log('Sheet indicator updated to:', sheetName);
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
 * Add a new task from the Add Task form.
 * Pure DOM/UI shell: reads inputs, delegates Excel work to createTask
 * (in tasks.js), then resets the form and shows a status banner.
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

  await withStatus('Add task', async () => {
    await Excel.run(async (context) => {
      const result = await createTask(context, name, weight);
      if (nameInput) nameInput.value = '';
      if (weightInput) weightInput.value = '1';
      showStatus(`✅ Task "${result.name}" added!`, 'success');
    });
  });
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

  const colors = {
    warning: '#e65100',
    error: '#c62828',
    success: '#2e7d32',
    info: '#1565c0',
  };
  const accent = colors[type] || colors.info;

  // Backdrop
  const modal = document.createElement('div');
  modal.id = 'custom-modal';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');
  modal.setAttribute('aria-labelledby', 'custom-modal-title');
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

  // Inner card
  const card = document.createElement('div');
  card.style.cssText = `
    background: white;
    border-radius: 12px;
    padding: 20px;
    max-width: 280px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    text-align: center;
  `;

  // Title — textContent prevents HTML injection via the title arg.
  const titleEl = document.createElement('div');
  titleEl.id = 'custom-modal-title';
  titleEl.style.cssText = `
    font-size: 18px;
    font-weight: bold;
    color: ${accent};
    margin-bottom: 12px;
  `;
  titleEl.textContent = title;

  // Message — textContent prevents HTML injection via the message arg.
  const msgEl = document.createElement('div');
  msgEl.style.cssText = `
    font-size: 14px;
    color: #333;
    margin-bottom: 16px;
    line-height: 1.4;
  `;
  msgEl.textContent = message;

  // OK button
  const okBtn = document.createElement('button');
  okBtn.id = 'modal-ok-btn';
  okBtn.type = 'button';
  okBtn.textContent = 'OK';
  okBtn.style.cssText = `
    background: ${accent};
    color: white;
    border: none;
    padding: 10px 30px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: bold;
    cursor: pointer;
  `;
  okBtn.addEventListener('click', () => modal.remove());

  card.appendChild(titleEl);
  card.appendChild(msgEl);
  card.appendChild(okBtn);
  modal.appendChild(card);
  document.body.appendChild(modal);

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
window.showWarningPopup = showWarningPopup;
window.showInfoPopup = showInfoPopup;
window.showModal = showModal;

/**
 * Run an async function and surface failures through showStatus.
 *
 * Removes the repetitive
 *   try { ... } catch (e) { showStatus('Error: ' + e.message, 'error') }
 * boilerplate from public entry points. The wrapped function is
 * responsible for emitting its own success status (if any) — this
 * wrapper only handles the error path.
 *
 * @param {string} label - Short label used in the error message,
 *   e.g. 'Sort habits'. The full error becomes
 *   '{label} failed: {error.message}'.
 * @param {() => Promise<T>} fn - Async function to run.
 * @returns {Promise<T | undefined>} The fn return value, or
 *   undefined if it threw.
 */
async function withStatus(label, fn) {
  try {
    return await fn();
  } catch (e) {
    console.error(`${label} failed:`, e);
    showStatus(`${label} failed: ${e.message}`, 'error');
    return undefined;
  }
}

window.withStatus = withStatus;
