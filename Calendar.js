/**
 * Google Calendar -> Weekly grid sync (Google Sheets edition).
 *
 * One-way import: reads the current week's events from every calendar in
 * the user's Google Calendar list (primary + shared/subscribed) and drops
 * each event's title into the matching day/time slot(s) of the Weekly
 * timetable. The grid is 30-minute rows from CONFIG.WEEKLY.FIRST_HOUR (8:00)
 * through 23:30.
 *
 * Which calendars: all of them by default. Set
 * CONFIG.CALENDAR.CALENDAR_NAMES to a list of names (or calendar ids) to
 * import only those, and CONFIG.CALENDAR.EXCLUDE_CALENDAR_NAMES to leave
 * some out (e.g. a work calendar, or a holiday feed).
 *
 * Overlaps: when two events land in the same 30-minute slot (common once
 * several calendars are imported), their titles are joined with
 * CONFIG.CALENDAR.MULTI_EVENT_SEPARATOR so nothing is silently overwritten.
 *
 * Re-syncable: the cells written by the previous sync are remembered in
 * DocumentProperties and cleared first (unless the user has since scored
 * them), so re-running updates in place instead of duplicating. Existing
 * cell contents in the synced range are overwritten by calendar events.
 *
 * Dropdowns: the task-cell dropdown is never removed for good. Event cells
 * are left without one while they hold a free-form title (so it isn't
 * flagged invalid), and the whole grid's dropdowns are re-applied at the
 * end of every sync and at week rollover.
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
 * Lower-case + trim a calendar name/id for case-insensitive comparison.
 * @param {*} name
 * @returns {string}
 */
function normalizeCalendarName_(name) {
  return String(name == null ? '' : name).trim().toLowerCase();
}

/**
 * Calendar names configured for import (CONFIG.CALENDAR.CALENDAR_NAMES,
 * falling back to the legacy single CALENDAR_NAME).
 * @returns {string[]} empty means "every calendar"
 */
function configuredCalendarNames_() {
  let raw = CONFIG.CALENDAR.CALENDAR_NAMES;
  // Tolerate a single name typed as a plain string instead of an array.
  if (typeof raw === 'string') raw = [raw];
  const names = [];
  if (Array.isArray(raw)) {
    for (let i = 0; i < raw.length; i++) {
      const n = String(raw[i] == null ? '' : raw[i]).trim();
      if (n) names.push(n);
    }
  }
  if (names.length) return names;
  const legacy = String(CONFIG.CALENDAR.CALENDAR_NAME || '').trim();
  return legacy ? [legacy] : [];
}

/**
 * Calendar names to leave out, normalized for comparison.
 * @returns {string[]}
 */
function excludedCalendarNames_() {
  const raw = CONFIG.CALENDAR.EXCLUDE_CALENDAR_NAMES;
  const out = [];
  if (Array.isArray(raw)) {
    for (let i = 0; i < raw.length; i++) {
      const n = normalizeCalendarName_(raw[i]);
      if (n) out.push(n);
    }
  }
  return out;
}

/**
 * @param {GoogleAppsScript.Calendar.Calendar} cal
 * @returns {string} the calendar's id, or '' when unavailable
 */
function calendarId_(cal) {
  try {
    return cal.getId() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Does this calendar match any of the given names (case-insensitive,
 * by display name or calendar id)?
 * @param {GoogleAppsScript.Calendar.Calendar} cal
 * @param {string[]} names
 * @returns {boolean}
 */
function calendarMatches_(cal, names) {
  const name = normalizeCalendarName_(cal.getName());
  const id = normalizeCalendarName_(calendarId_(cal));
  for (let i = 0; i < names.length; i++) {
    const wanted = normalizeCalendarName_(names[i]);
    if (!wanted) continue;
    if (wanted === name) return true;
    if (id && wanted === id) return true;
  }
  return false;
}

/**
 * Honor CONFIG.CALENDAR.ONLY_SELECTED when the runtime exposes
 * Calendar.isSelected(); older runtimes are treated as "selected".
 * @param {GoogleAppsScript.Calendar.Calendar} cal
 * @returns {boolean}
 */
function isSelectedCalendar_(cal) {
  if (!CONFIG.CALENDAR.ONLY_SELECTED) return true;
  try {
    if (typeof cal.isSelected === 'function') return cal.isSelected() !== false;
  } catch (e) {
    /* fall through: never drop a calendar just because the check failed */
  }
  return true;
}

/**
 * Is this calendar named in an already-normalized exclusion list?
 * @param {GoogleAppsScript.Calendar.Calendar} cal
 * @param {string[]} excluded normalized names/ids
 * @returns {boolean}
 */
function calendarIsExcluded_(cal, excluded) {
  if (!excluded.length) return false;
  const name = normalizeCalendarName_(cal.getName());
  const id = normalizeCalendarName_(calendarId_(cal));
  for (let i = 0; i < excluded.length; i++) {
    if (excluded[i] === name) return true;
    if (id && excluded[i] === id) return true;
  }
  return false;
}

/**
 * Resolve which calendars to read from: every calendar in the user's list
 * minus CONFIG.CALENDAR.EXCLUDE_CALENDAR_NAMES, or only the names listed in
 * CONFIG.CALENDAR.CALENDAR_NAMES when that list is non-empty.
 * @returns {{calendars: GoogleAppsScript.Calendar.Calendar[], missing: string[]}}
 *   `missing` lists configured names that matched no calendar.
 */
function getSourceCalendars_() {
  const all = CalendarApp.getAllCalendars() || [];
  const wanted = configuredCalendarNames_();
  const excluded = excludedCalendarNames_();
  const missing = wanted.slice();
  const calendars = [];
  const seenIds = {};

  for (let i = 0; i < all.length; i++) {
    const cal = all[i];

    // De-dupe: the same calendar can be listed twice in rare setups
    // (e.g. subscribed by name and by id).
    const id = calendarId_(cal);
    if (id && seenIds[id]) continue;

    if (wanted.length && !calendarMatches_(cal, wanted)) continue;

    // A configured name was found — stop reporting it as missing.
    if (wanted.length) {
      for (let m = missing.length - 1; m >= 0; m--) {
        if (calendarMatches_(cal, [missing[m]])) missing.splice(m, 1);
      }
    }

    if (!isSelectedCalendar_(cal)) continue;
    if (calendarIsExcluded_(cal, excluded)) continue;

    if (id) seenIds[id] = true;
    calendars.push(cal);
  }

  return { calendars: calendars, missing: missing };
}

/**
 * Gather each calendar's events for the week, isolating per-calendar
 * failures (a calendar we cannot read must not abort the whole sync).
 * @param {GoogleAppsScript.Calendar.Calendar[]} calendars
 * @param {Date} weekStart
 * @param {Date} weekEnd exclusive
 * @param {string[]} unreadable collects "name (error)" entries
 * @returns {GoogleAppsScript.Calendar.CalendarEvent[]}
 */
function collectWeekEvents_(calendars, weekStart, weekEnd, unreadable) {
  const out = [];
  for (let c = 0; c < calendars.length; c++) {
    const cal = calendars[c];
    let events;
    try {
      events = cal.getEvents(weekStart, weekEnd) || [];
    } catch (e) {
      unreadable.push(cal.getName() + ' (' + (e && e.message ? e.message : e) + ')');
      continue;
    }
    for (let i = 0; i < events.length; i++) out.push(events[i]);
  }
  return out;
}

/**
 * The signed-in account's email, derived from the primary calendar's id
 * (a Gmail/Workspace primary calendar is identified by the account address).
 * This avoids needing the userinfo.email scope just to inspect guest lists.
 * @returns {string} lower-cased email, or '' when it cannot be determined
 */
function primaryCalendarEmail_() {
  try {
    return String(CalendarApp.getDefaultCalendar().getId() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * Did the user actually decline this event? getMyStatus() can report NO for
 * events on shared/subscribed calendars that the user was never invited to,
 * so "NO" alone is not enough — the user must also be on the guest list.
 * When the account email is unknown the event is kept (never drop silently).
 * @param {GoogleAppsScript.Calendar.CalendarEvent} ev
 * @param {string} accountEmail lower-cased primary calendar email
 * @returns {boolean}
 */
function isDeclinedEvent_(ev, accountEmail) {
  if (!accountEmail) return false;
  try {
    if (ev.getMyStatus() !== CalendarApp.GuestStatus.NO) return false;
    const guests = ev.getGuestList() || [];
    for (let i = 0; i < guests.length; i++) {
      let email = '';
      try {
        email = String(guests[i].getEmail() || '').trim().toLowerCase();
      } catch (e) {
        email = '';
      }
      if (email && email === accountEmail) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Stable identity for an event occurrence, used to avoid importing the same
 * event twice when it is visible through more than one calendar.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} ev
 * @param {Date} start
 * @returns {string}
 */
function eventDedupeKey_(ev, start) {
  let id = '';
  try {
    id = ev.getId() || '';
  } catch (e) {
    id = '';
  }
  if (id) return id + '|' + start.getTime();
  return ev.getTitle() + '|' + start.getTime();
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
    // Content only — the task dropdown is restored grid-wide at the end of
    // the sync (applyTaskValidation_), so clearing must not remove it.
    if (score === '' || score === null) cell.clearContent();
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

  const sources = getSourceCalendars_();
  if (!sources.calendars.length) {
    let msg = 'No calendar found to sync.';
    if (sources.missing.length) {
      msg += ' Not found in your Calendar list: ' + sources.missing.join(', ') + '.';
    } else if (CONFIG.CALENDAR.ONLY_SELECTED || excludedCalendarNames_().length) {
      msg += ' Every calendar was filtered out (ONLY_SELECTED / EXCLUDE_CALENDAR_NAMES).';
    }
    return msg;
  }

  const W = CONFIG.WEEKLY;
  const weekStart = new Date(monday);
  const weekEnd = new Date(monday);
  weekEnd.setDate(monday.getDate() + W.DAYS_IN_WEEK); // exclusive: next Monday 00:00

  const unreadable = [];
  const rawEvents = collectWeekEvents_(sources.calendars, weekStart, weekEnd, unreadable);
  const slotCount = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;
  const separator = CONFIG.CALENDAR.MULTI_EVENT_SEPARATOR;

  // Normalize + filter every event into grid-ready items. Filtering happens
  // before sorting so the overlap join order is stable.
  const items = [];
  const seenEvents = {};
  const accountEmail = CONFIG.CALENDAR.SKIP_DECLINED ? primaryCalendarEmail_() : '';
  let allDaySkipped = 0;
  let declinedSkipped = 0;

  for (let e = 0; e < rawEvents.length; e++) {
    const ev = rawEvents[e];
    if (CONFIG.CALENDAR.SKIP_ALL_DAY && ev.isAllDayEvent()) {
      allDaySkipped++;
      continue;
    }
    const title = ev.getTitle();
    if (!title) continue;
    if (CONFIG.CALENDAR.SKIP_DECLINED && isDeclinedEvent_(ev, accountEmail)) {
      declinedSkipped++;
      continue;
    }

    const start = ev.getStartTime();
    const end = ev.getEndTime();
    if (!start || !end) continue;

    const startMidnight = new Date(start);
    startMidnight.setHours(0, 0, 0, 0);
    const dayIndex = daysBetween(monday, startMidnight);
    if (dayIndex < 0 || dayIndex >= W.DAYS_IN_WEEK) continue;

    const key = eventDedupeKey_(ev, start);
    if (seenEvents[key]) continue;
    seenEvents[key] = true;

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

    items.push({
      dayIndex: dayIndex,
      startDec: startDec,
      endDec: endDec,
      title: title,
      startMs: start.getTime(),
      durationMs: end.getTime() - start.getTime(),
    });
  }

  // Earlier events first; longer events win ties, then title, so the joined
  // cell content is deterministic no matter which calendar answered first.
  items.sort(function (a, b) {
    if (a.startMs !== b.startMs) return a.startMs - b.startMs;
    if (a.durationMs !== b.durationMs) return b.durationMs - a.durationMs;
    return a.title < b.title ? -1 : a.title > b.title ? 1 : 0;
  });

  // Clear last sync's cells first so removed/moved events don't linger.
  clearPreviousSyncCells_(sheet);

  // One bucket per grid cell: several overlapping events share a single
  // cell instead of the last one overwriting the others.
  const buckets = {};
  const written = [];
  let eventCount = 0;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const taskCol = getTaskColForDay(item.dayIndex);
    const slots = eventSlotIndices_(item.startDec, item.endDec, W.FIRST_HOUR, slotCount);
    if (slots.length === 0) continue;
    eventCount++;

    for (let s = 0; s < slots.length; s++) {
      const row = W.DATA_START_ROW + slots[s];
      const cellKey = row + ',' + taskCol;
      let bucket = buckets[cellKey];
      if (!bucket) {
        bucket = buckets[cellKey] = { row: row, col: taskCol, titles: [] };
        written.push(cellKey);
      }
      if (bucket.titles.indexOf(item.title) < 0) bucket.titles.push(item.title);
    }
  }

  for (let w = 0; w < written.length; w++) {
    const bucket = buckets[written[w]];
    sheet.getRange(bucket.row, bucket.col).setValue(bucket.titles.join(separator));
  }

  saveSyncCells_(written);

  // Restore the task dropdown grid-wide. Cells holding this week's event
  // titles keep none (so the free-form title isn't flagged invalid — the
  // default; config-gated), but every other task cell is guaranteed the
  // dropdown again. Without this, a timeslot that once held an event lost
  // its dropdown permanently, including for next week's grid.
  safeInit_('Task dropdown refresh failed', function () {
    applyTaskValidation_(sheet, CONFIG.CALENDAR.CLEAR_DROPDOWN !== false);
  });

  SpreadsheetApp.flush();

  let msg = 'Synced ' + eventCount + ' event(s) from ' + sources.calendars.length +
    ' calendar(s) into ' + written.length + ' slot(s) for the week of ' +
    monday.toDateString() + '.';
  const skipped = [];
  if (allDaySkipped) skipped.push(allDaySkipped + ' all-day');
  if (declinedSkipped) skipped.push(declinedSkipped + ' declined');
  if (skipped.length) msg += ' Skipped ' + skipped.join(' + ') + ' event(s).';
  if (sources.missing.length) {
    msg += ' Calendar(s) not found: ' + sources.missing.join(', ') + '.';
  }
  if (unreadable.length) {
    msg += ' Could not read: ' + unreadable.join('; ') + '.';
  }
  return msg;
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
