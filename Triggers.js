/**
 * Triggers, menu, and sidebar entry points (Google Sheets edition).
 *
 * Replaces src/taskpane/js/app.js + events.js. Mapping:
 *   - Office.onReady / Workbook_Open  -> onOpen(e) (menu + sidebar)
 *   - sheet.onChanged                 -> handleEdit(e) (installable onEdit)
 *   - Application.OnTime (unsupported) -> dailyMaintenance (time-driven)
 *   - app.js#initializeAddin          -> initialize() (called from sidebar)
 *
 * The edit handler is named handleEdit (not onEdit) on purpose: it is
 * installed as an *installable* trigger so it runs with full
 * authorization (LockService, Summary writes). A function literally
 * named onEdit would also fire as a limited simple trigger and double-run.
 */

// ----------------------------------------------------------------------
// Open: menu + auto-sidebar
// ----------------------------------------------------------------------

/**
 * Simple trigger: runs on every open. Builds the menu and opens the
 * sidebar. The heavy, authorized init runs from the sidebar via
 * google.script.run.initialize().
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Weekly Plan')
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Set up sheets', 'setUpSheets')
    .addItem('Install triggers (run once)', 'installTriggers')
    .addSeparator()
    .addItem('Run daily init', 'initialize')
    .addItem('Recalculate now', 'recalculateWeek')
    .addItem('Add task…', 'addTaskPrompt_')
    .addItem('Random pick', 'randomPickFromUI')
    .addItem('Sync calendar → Weekly', 'syncCalendarFromUI')
    .addSeparator()
    .addItem('Export all to Drive', 'exportAllFromUI')
    .addItem('Send Telegram recap now', 'sendTelegramNowFromUI')
    .addItem('Send meal story now', 'sendMealStoryNowFromUI')
    .addItem('Set up Telegram…', 'setUpTelegram')
    .addItem('Set up Gemini…', 'setUpGemini')
    .addItem('Set up Diary…', 'setUpDiary')
    .addItem('Send diary reminder now', 'sendDiaryReminderFromUI')
    .addItem('Email summary now', 'sendSummaryEmailNowFromUI')
    .addItem('Start new week…', 'startNewWeekWithConfirm_')
    .addToUi();

  showSidebar();
}

/**
 * Render the sidebar from Sidebar.html.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Weekly Plan')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ----------------------------------------------------------------------
// Edit dispatch (installable onEdit trigger -> handleEdit)
// ----------------------------------------------------------------------

/**
 * Installable onEdit handler. Routes Weekly score edits and Habits
 * checkbox edits to their domain functions.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const name = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (name === CONFIG.WEEKLY_SHEET) {
      const W = CONFIG.WEEKLY;
      if (
        W.SCORE_COLUMNS.indexOf(col) !== -1 &&
        row >= W.DATA_START_ROW &&
        row <= W.LAST_TIME_ROW
      ) {
        // Guard against re-scoring: once a score cell has a value,
        // editing it again would double-count the contribution without
        // subtracting the old one. The oldValue property is set by the
        // installable onEdit trigger for single-cell edits.
        if (e.oldValue !== undefined && e.oldValue !== '' && e.oldValue !== null) {
          toast_('Score already set — cannot modify once scored.', 'Weekly Plan', 'warning');
          return;
        }
        const score = parseFloat(e.range.getValue());
        if (!isNaN(score)) processWeeklyScoreChange(row, col, score);
      }

      // Task cell edited — if the name matches a habit/task that has a
      // hyperlink, copy the link so the Weekly cell is also clickable.
      if (
        W.TASK_COLUMNS.indexOf(col) !== -1 &&
        row >= W.DATA_START_ROW &&
        row <= W.LAST_TIME_ROW
      ) {
        copyLinkToWeeklyCell_(e.range);
      }
      return;
    }

    if (name === CONFIG.HABITS_SHEET) {
      const H = CONFIG.HABITS;
      const checkboxCol = columnLetterToIndex(H.COLUMNS.DONE_CHECKBOX) + 1;
      if (col === checkboxCol && row >= H.DATA_START_ROW && e.range.getValue() === true) {
        recordHabitDone(row);
      }
      return;
    }
  } catch (err) {
    Logger.log('handleEdit error: ' + (err && err.message ? err.message : err));
  }
}

// ----------------------------------------------------------------------
// Selection dispatch (installable onSelectionChange -> handleSelection)
// ----------------------------------------------------------------------

/**
 * Simple trigger: fires on every selection change. Powers the clickable
 * control buttons in the Weekly sheet's control row.
 *
 * NOTE: onSelectionChange exists ONLY as a simple trigger — there is no
 * installable equivalent, so this function must keep this exact name and
 * is NOT created in installTriggers. Simple triggers run without granted
 * scopes, so getUi() may be unavailable; UI actions fall back to a toast.
 * @param {GoogleAppsScript.Events.SheetsOnSelectionChange} e
 */
function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.WEEKLY_SHEET) return;
    if (e.range.getRow() !== CONFIG.WEEKLY.CONTROL_ROW) return;

    const col = e.range.getColumn();
    const buttons = CONFIG.WEEKLY.CONTROL_BUTTONS;
    let action = null;
    for (let i = 0; i < buttons.length; i++) {
      if (buttons[i].col === col) {
        action = buttons[i].action;
        break;
      }
    }
    if (!action) return;

    const ui = getUiOrNull_();
    switch (action) {
      case 'random':
        randomPick();
        break;
      case 'help':
        if (ui) showHelpDialog_();
        else toast_('Tasks → columns C,E,G…; scores → next column D,F,H….', 'Help');
        break;
      case 'add':
        if (ui) addTaskPrompt_();
        else toast_('Open the sidebar → Add Task.', 'Add Task');
        break;
      case 'delete':
        if (ui) deleteTaskPrompt_();
        else toast_('Delete a task from the Tasks sheet or the menu.', 'Delete Task');
        break;
      case 'thanks':
        if (ui) {
          ui.alert(
            'Thanks!',
            'Weekly Plan — migrated from VBA → Office.js → Google Sheets. Enjoy your week!',
            ui.ButtonSet.OK
          );
        } else {
          toast_('Thanks for using Weekly Plan!', 'Thanks');
        }
        break;
    }
  } catch (err) {
    Logger.log('onSelectionChange error: ' + (err && err.message ? err.message : err));
  }
}

/**
 * Return the spreadsheet Ui, or null if unavailable (e.g. a simple
 * trigger running without the container.ui scope).
 * @returns {GoogleAppsScript.Base.Ui|null}
 */
function getUiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

/**
 * Show a brief help dialog for the Weekly sheet.
 */
function showHelpDialog_() {
  SpreadsheetApp.getUi().alert(
    'How to use',
    'Tasks: type a task in a task column (C, E, G…).\n' +
      'Scores: pick 0–1 in the score column to its right (D, F, H…).\n' +
      'Random Fill: fills empty current-day slots.\n' +
      'Habits: tick the checkbox next to a habit on the Habits sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Delete a task via a native prompt (control-bar "Delete Task").
 */
function deleteTaskPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Delete Task', 'Task name to delete:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const name = resp.getResponseText().trim();
  if (!name) return;
  const ok = deleteTask(name);
  ui.alert(ok ? 'Deleted task "' + name + '".' : 'No task named "' + name + '" found.');
}

// ----------------------------------------------------------------------
// Trigger installation
// ----------------------------------------------------------------------

/**
 * One-time, idempotent installer for the installable onEdit trigger and
 * the daily time-driven maintenance trigger. Removes existing copies
 * first so re-running never stacks duplicates.
 * @returns {string} status message
 */
function installTriggers() {
  const managed = ['handleEdit', 'dailyMaintenance', 'morningTelegram', 'mealStory', 'refreshTimeHighlight', 'diaryReminder'];
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (managed.indexOf(triggers[i].getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('handleEdit').forSpreadsheet(getSpreadsheet_()).onEdit().create();
  ScriptApp.newTrigger('dailyMaintenance').timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger('morningTelegram')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TELEGRAM.SEND_HOUR)
    .create();

  // One meal-story trigger per configured meal hour (breakfast/lunch/dinner).
  const mealHours = CONFIG.STORY.MEAL_HOURS || [];
  for (let h = 0; h < mealHours.length; h++) {
    ScriptApp.newTrigger('mealStory').timeBased().everyDays(1).atHour(mealHours[h]).create();
  }

  // Update the current-time-row highlight every 5 min so it stays accurate
  // even when the sidebar is closed.  5 min keeps worst-case lag negligible
  // for 30-min time slots (~288 executions/day, well within quota).
  ScriptApp.newTrigger('refreshTimeHighlight').timeBased().everyMinutes(5).create();

  // Evening diary reminder (~22:30, ±15 min window). The onFormSubmit
  // trigger for handleDiarySubmit is installed separately by setUpDiary.
  ScriptApp.newTrigger('diaryReminder')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.DIARY.SEND_HOUR)
    .nearMinute(CONFIG.DIARY.SEND_MINUTE)
    .create();

  // onSelectionChange (control buttons) is a SIMPLE trigger — it fires
  // automatically by name and cannot/should not be created here.
  const msg =
    'Triggers installed: live edits + daily maintenance (~5am) + Telegram recap (~' +
    CONFIG.TELEGRAM.SEND_HOUR + 'am) + meal stories (' +
    mealHours.join(', ') + ') + diary reminder (~' +
    CONFIG.DIARY.SEND_HOUR + ':' + String(CONFIG.DIARY.SEND_MINUTE).padStart(2, '0') +
    '). Control buttons work automatically.';
  toast_(msg, 'Weekly Plan');
  return msg;
}

// ----------------------------------------------------------------------
// Initialization (open-time + daily)
// ----------------------------------------------------------------------

/**
 * Shared init core: new-day/new-week handling + highlights. Safe to run
 * with or without an active UI.
 */
function runDailyInitCore_() {
  safeInit_('Habits init skipped', function () {
    if (getSheetByName_(CONFIG.HABITS_SHEET)) initializeHabitsSheet();
  });
  safeInit_('Weekly init skipped', function () {
    if (getSheetByName_(CONFIG.WEEKLY_SHEET)) initializeWeeklyOnOpen();
  });
  safeInit_('Daily quest skipped', function () {
    ensureDailyQuest_();
  });
  safeInit_('Weekly boss skipped', function () {
    ensureWeeklyBoss_();
  });
  // Diary: refresh the top band (self-heals after a rollover clears it) and
  // apply deferred habit/badge side effects in this spreadsheet context.
  safeInit_('Diary band refresh skipped', function () {
    refreshWeeklyFocusBand_();
  });
  safeInit_('Diary habits skipped', function () {
    processPendingDiaryHabits_();
  });
  safeInit_('Diary chronicler badge skipped', function () {
    maybeAwardChronicler_();
  });
  PropertiesService.getDocumentProperties().setProperty(
    'lastInitDate',
    formatDateYYYYMMDD(new Date())
  );
}

/**
 * Called from the sidebar on load (full authorization). Activates the
 * Weekly sheet, runs init, and reports the current sheet name.
 * @returns {{sheet: string}}
 */
function initialize() {
  const weekly = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (weekly) safeInit_('Activate Weekly failed', function () { weekly.activate(); });
  runDailyInitCore_();
  return { sheet: getCurrentSheetName() };
}

/**
 * Time-driven daily trigger entry point (~5am). Runs init even when no
 * one has the sheet open, so archiving/reset and highlights stay current,
 * and syncs the calendar. The morning recap itself is sent separately by
 * the `morningTelegram` trigger (~CONFIG.TELEGRAM.SEND_HOUR) so it lands
 * when the user actually starts their day.
 */
function dailyMaintenance() {
  runDailyInitCore_();
  safeInit_('Calendar sync skipped', function () {
    syncCalendarToWeekly_();
  });
}

/**
 * Time-driven morning trigger (~CONFIG.TELEGRAM.SEND_HOUR). Sends the
 * Telegram morning recap once per day (deduped + enable-gated inside
 * sendMorningTelegram_). The quest self-heals via ensureDailyQuest_, so
 * this stands on its own even if dailyMaintenance hasn't run yet.
 */
function morningTelegram() {
  safeInit_('Morning Telegram skipped', function () {
    sendMorningTelegram_(false);
  });
}

/**
 * Time-driven meal trigger (one per CONFIG.STORY.MEAL_HOURS). Picks a
 * random habit, has Gemini write a short motivating story about it, and
 * sends it to Telegram. Deduped + enable-gated inside sendMealStory_.
 */
function mealStory() {
  safeInit_('Meal story skipped', function () {
    sendMealStory_(false);
  });
}

/**
 * Time-driven evening trigger (~22:30). Sends the prefilled diary form link
 * to Telegram once per day (deduped + enable-gated inside sendDiaryReminder_).
 */
function diaryReminder() {
  safeInit_('Diary reminder skipped', function () {
    sendDiaryReminder_(false);
  });
}

// ----------------------------------------------------------------------
// Sidebar / menu-callable wrappers (return serializable values)
// ----------------------------------------------------------------------

/** @returns {string} active sheet name */
function getCurrentSheetName() {
  return getSpreadsheet_().getActiveSheet().getName();
}

/** @returns {{positive:number,negative:number,total:number}|null} */
function getTodayScoreFromUI() {
  return getTodayScore();
}

/**
 * Recent activity-log entries for the sidebar "Activity Log" panel
 * (newest first). Captures every toast_ — including those fired from
 * edit/selection triggers that the sidebar can't otherwise observe.
 * @returns {Array<{ts:number,title:string,msg:string,type:string}>}
 */
function getActivityLogFromUI() {
  return getActivityLog_();
}

/** @returns {number} slots filled */
function randomPickFromUI() {
  return randomPick();
}

/** @returns {string} status message */
function recalculateWeekFromUI() {
  return recalculateWeek();
}

/**
 * Add a task from the sidebar form.
 * @param {string} name
 * @param {number} weight
 * @returns {{row:number,name:string,weight:number}}
 */
function addTaskFromUI(name, weight) {
  const result = createTask(name, weight);
  toast_('Task "' + result.name + '" added.', 'Weekly Plan');
  return result;
}

/**
 * Add a task via a native prompt (menu path, no sidebar needed).
 */
function addTaskPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const nameResp = ui.prompt('Add Task', 'Task name:', ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResp.getResponseText().trim();
  if (!name) {
    ui.alert('Task name is required.');
    return;
  }
  const wResp = ui.prompt('Add Task', 'Weight (score multiplier, default 1):', ui.ButtonSet.OK_CANCEL);
  if (wResp.getSelectedButton() !== ui.Button.OK) return;
  const weight = parseFloat(wResp.getResponseText()) || 1;
  const result = createTask(name, weight);
  ui.alert('Task "' + result.name + '" added (weight ' + result.weight + ').');
}

/** @returns {string} status message */
function sortHabitsFromUI() {
  return sortHabits();
}

/** @returns {string} status message */
function sortTasksFromUI() {
  return sortTasks();
}

/** @returns {string} status message */
function refreshHabitDatesFromUI() {
  return refreshHabitsDates();
}

/** @returns {{count:number, folderUrl:string|null}} */
function exportAllFromUI() {
  return exportAllSheets();
}

/** @returns {string|null} Drive file URL */
function exportSummaryFromUI() {
  return exportSummaryData();
}

/**
 * Refresh the Weekly time-row highlight.  Called by the sidebar's 60 s
 * poller AND by the background 30-min time-driven trigger so the highlight
 * stays current even when the sidebar is closed.
 */
function refreshTimeHighlight() {
  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (sheet) highlightCurrentTimeRow(sheet);
}

/**
 * Sidebar wrapper — delegates to the shared refreshTimeHighlight above.
 * @deprecated Kept for backward compat; the sidebar can call refreshTimeHighlight directly.
 */
function refreshTimeHighlightFromUI() {
  refreshTimeHighlight();
}

/** @returns {string} status message */
function setUpSheetsFromUI() {
  return setUpSheets();
}

/**
 * Start a new week, gated behind a confirmation dialog (the destructive
 * reset has no confirm in the Office build).
 * @returns {string} status message
 */
function startNewWeekWithConfirm_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Start a new week?',
    'This archives the current week (Archive sheet + Drive CSV) and clears the grid. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return 'Cancelled.';

  const result = doWeekRollover_(true);
  const msg =
    'New week started. Archived ' +
    result.rows +
    ' row(s)' +
    (result.url ? ' — CSV: ' + result.url : '') +
    '.';
  toast_('New week started.', 'Weekly Plan');
  return msg;
}

/**
 * Sidebar-callable new-week (shows the same confirm dialog).
 * @returns {string} status message
 */
function startNewWeekFromUI() {
  return startNewWeekWithConfirm_();
}
