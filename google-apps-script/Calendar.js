/**
 * Google Calendar -> Weekly grid sync (Google Sheets edition).
 *
 * One-way import: reads the current week's events from the primary (or a
 * named) calendar and drops each event's title into the matching day/time
 * slot(s) of the Weekly timetable. The grid is 30-minute rows from
 * CONFIG.WEEKLY.FIRST_HOUR (8:00) through 23:30.
 *
 * Re-syncable: the cells written by the previous sync are remembered in
 * DocumentProperties and cleared first (unless the user has since scored
 * them), so re-running updates in place instead of duplicating. Existing
 * cell contents in the synced range are overwritten by calendar events.
 *
 * Requires the calendar.readonly OAuth scope (see appsscript.json).
 */

/**
 * Slot indices a time span [startDec, endDec) overlaps, within a grid of
 * `slotCount` 30-minute rows starting at `firstHour`. Pure + testable.
 * @param {number} startDec start time in decimal hours (e.g. 9.5 = 09:30)
 * @param {number} endDec end time in decimal hours
 * @param {number} firstHour grid's first hour (e.g. 8)
 * @param {number} slotCount number of 30-min rows
 * @returns {number[]} 0-based slot indices the span covers
 */
function eventSlotIndices_(startDec, endDec, firstHour, slotCount) {
  const out = [];
  for (let i = 0; i < slotCount; i++) {
    const slotStart = firstHour + i * 0.5;
    const slotEnd = slotStart + 0.5;
    if (startDec < slotEnd && endDec > slotStart) out.push(i);
  }
  return out;
}

/**
 * Resolve the calendar to read from: a configured name, else the primary.
 * @returns {GoogleAppsScript.Calendar.Calendar|null}
 */
function getSourceCalendar_() {
  const name = (CONFIG.CALENDAR.CALENDAR_NAME || '').trim();
  if (name) {
    const cals = CalendarApp.getCalendarsByName(name);
    return cals && cals.length ? cals[0] : null;
  }
  return CalendarApp.getDefaultCalendar();
}

/**
 * Clear cells written by the previous sync, skipping any the user has
 * since scored (a non-empty score column means they acted on it).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearPreviousSyncCells_(sheet) {
  const raw = PropertiesService.getDocumentProperties().getProperty(
    CONFIG.CALENDAR.SYNCED_CELLS_PROP
  );
  if (!raw) return;
  let cells = [];
  try {
    cells = JSON.parse(raw) || [];
  } catch (e) {
    cells = [];
  }
  for (let i = 0; i < cells.length; i++) {
    const parts = String(cells[i]).split(',');
    if (parts.length !== 2) continue;
    const r = parseInt(parts[0], 10);
    const c = parseInt(parts[1], 10);
    if (!(r > 0 && c > 0)) continue;
    const cell = sheet.getRange(r, c);
    // Preserve cells the user already scored (score col = task col + 1).
    const score = sheet.getRange(r, c + 1).getValue();
    if (score === '' || score === null) cell.clearContent();
    // Strip the dropdown left over from older syncs so stale cells stop
    // showing "Input must fall within specified range".
    if (CONFIG.CALENDAR.CLEAR_DROPDOWN !== false) cell.setDataValidation(null);
  }
}

/**
 * Persist the list of "row,col" cells written by this sync.
 * @param {string[]} cells
 */
function saveSyncCells_(cells) {
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG.CALENDAR.SYNCED_CELLS_PROP,
    JSON.stringify(cells)
  );
}

/**
 * Import the current week's calendar events into the Weekly grid.
 * @returns {string} status message
 */
function syncCalendarToWeekly_() {
  if (!CONFIG.CALENDAR.ENABLED) return 'Calendar sync is disabled (Config.CALENDAR.ENABLED).';

  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) return 'Weekly sheet not found.';

  let monday = getSheetWeekMonday_(sheet);
  if (!monday) {
    // Week dates not initialized yet — set this week's dates, then retry.
    setNewWeekDates(sheet);
    monday = getSheetWeekMonday_(sheet);
  }
  if (!monday) return 'Week dates not set on the Weekly sheet yet. Run "Set up sheets" first.';

  const cal = getSourceCalendar_();
  if (!cal) return 'Source calendar not found.';

  const W = CONFIG.WEEKLY;
  const weekStart = new Date(monday);
  const weekEnd = new Date(monday);
  weekEnd.setDate(monday.getDate() + W.DAYS_IN_WEEK); // exclusive: next Monday 00:00

  const events = cal.getEvents(weekStart, weekEnd);
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;

  // Clear last sync's cells first so removed/moved events don't linger.
  clearPreviousSyncCells_(sheet);

  const written = [];
  let eventCount = 0;
  for (let e = 0; e < events.length; e++) {
    const ev = events[e];
    if (CONFIG.CALENDAR.SKIP_ALL_DAY && ev.isAllDayEvent()) continue;
    const title = ev.getTitle();
    if (!title) continue;

    const start = ev.getStartTime();
    const end = ev.getEndTime();

    const startMidnight = new Date(start);
    startMidnight.setHours(0, 0, 0, 0);
    const dayIndex = daysBetween(monday, startMidnight);
    if (dayIndex < 0 || dayIndex >= W.DAYS_IN_WEEK) continue;

    let startDec = start.getHours() + start.getMinutes() / 60;
    const endMidnight = new Date(end);
    endMidnight.setHours(0, 0, 0, 0);
    // End on a later day => clamp to end of the grid day.
    let endDec = daysBetween(startMidnight, endMidnight) >= 1 ? 24 : end.getHours() + end.getMinutes() / 60;
    if (endDec <= startDec) endDec = startDec + 0.5; // zero/negative length -> one slot

    if (endDec <= W.FIRST_HOUR) continue; // entirely before the grid window
    if (startDec >= 24) continue; // entirely after the grid window
    startDec = Math.max(startDec, W.FIRST_HOUR);
    endDec = Math.min(endDec, 24);

    const taskCol = getTaskColForDay(dayIndex);
    const slots = eventSlotIndices_(startDec, endDec, W.FIRST_HOUR, slotCount);
    if (slots.length === 0) continue;

    for (let s = 0; s < slots.length; s++) {
      const row = W.DATA_START_ROW + slots[s];
      const cell = sheet.getRange(row, taskCol);
      cell.setValue(title);
      // Synced titles are free-form, not Tasks-list entries — drop the
      // dropdown so they aren't flagged invalid (default on).
      if (CONFIG.CALENDAR.CLEAR_DROPDOWN !== false) cell.setDataValidation(null);
      written.push(row + ',' + taskCol);
    }
    eventCount++;
  }

  saveSyncCells_(written);
  SpreadsheetApp.flush();

  return 'Synced ' + eventCount + ' event(s) into ' + written.length + ' slot(s) for the week of ' +
    monday.toDateString() + '.';
}

/**
 * Sidebar/menu-callable calendar sync (toasts + returns status).
 * @returns {string} status message
 */
function syncCalendarFromUI() {
  const msg = syncCalendarToWeekly_();
  toast_(msg, 'Weekly Plan');
  return msg;
}
