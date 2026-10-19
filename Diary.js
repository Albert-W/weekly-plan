/**
 * Daily Diary feature (Google Sheets edition).
 *
 * Morning plan, not an evening summary: the ~8:30 Telegram message (see
 * Telegram.js) links a prefilled Google Form, and submitting it upserts one
 * row in the Diary sheet (keyed by date), refreshes the Weekly top band
 * ("📌 今日要点" on row 1, "🎯 明日计划" on row 3), and marks the "写日记"
 * habit done — deferred to a spreadsheet-context run.
 *
 * There are two entry points onto the same row: this form, and the in-page
 * editor on the local reader (`saveDiaryFromWeb_`, Web App POST action=diary).
 * Both call upsertDiaryRow_.
 *
 * Why the habit + badge are deferred (not marked inside handleDiarySubmit):
 * the onFormSubmit trigger fires in a Form (cross-document) context, where
 * getDocumentProperties()/getDocumentLock() may resolve to the Form instead of
 * the bound spreadsheet. All gamification side effects (XP, combo, boss,
 * badges) are therefore applied by processPendingDiaryHabits_ / a
 * spreadsheet-context trigger, where they are safe. The Diary sheet write and
 * the band refresh only touch cells, so they run directly in the trigger.
 *
 * Persistent diary state lives in Script Properties (CONFIG.DIARY.*_PROP) for
 * the same cross-document reason — see the registry note at the top of Config.js.
 */

// ----------------------------------------------------------------------
// Setup
// ----------------------------------------------------------------------

/**
 * Create (or reuse) the diary Google Form, install the onFormSubmit trigger,
 * and ensure the `spreadsheetId` Script Property is set. Idempotent.
 * @returns {string} status message
 */
function setUpDiary() {
  const props = PropertiesService.getScriptProperties();
  const D = CONFIG.DIARY;

  // The form trigger has no active spreadsheet; the Diary sheet writes rely
  // on getSpreadsheet_()'s openById fallback keyed by this property.
  if (!props.getProperty(CONFIG.SYNC.SPREADSHEET_ID_PROP)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) props.setProperty(CONFIG.SYNC.SPREADSHEET_ID_PROP, ss.getId());
    } catch (e) {
      // Non-UI context — leave unset; setup normally runs from the menu.
    }
  }

  let form = null;
  const formId = props.getProperty(D.FORM_ID_PROP);
  if (formId) {
    try {
      form = FormApp.openById(formId);
    } catch (e) {
      form = null;
    }
  }

  if (!form) {
    form = FormApp.create('每日日记');
    props.setProperty(D.FORM_ID_PROP, form.getId());

    // Fixed order + titles. handleDiarySubmit matches by TITLE (not index), so
    // reordering items later won't misparse responses.
    form.addDateItem().setTitle(D.TITLES.DATE);
    form.addMultipleChoiceItem().setTitle(D.TITLES.MOOD).setChoiceValues(D.MOOD_EMOJI);
    form.addParagraphTextItem().setTitle(D.TITLES.KEY_POINTS);
    form.addParagraphTextItem().setTitle(D.TITLES.HIGHLIGHT);
    form.addParagraphTextItem().setTitle(D.TITLES.PLAN);
    form.addParagraphTextItem().setTitle(D.TITLES.IMPROVE);
  }

  // Rename legacy item titles in place (e.g. '担忧' → '今日要点'), so a live
  // form migrates instead of ending up with a stale question that
  // handleDiarySubmit no longer looks for.
  const renames = D.RENAMED_TITLES || {};
  form.getItems().forEach(function (it) {
    const oldTitle = String(it.getTitle());
    if (renames[oldTitle]) it.setTitle(renames[oldTitle]);
  });

  // Idempotent for existing forms: append any new items that a code update
  // introduced (matched by title, like handleDiarySubmit). Running "Set up
  // Diary…" again on a live form picks up newly added questions.
  const have = {};
  form.getItems().forEach(function (it) { have[String(it.getTitle())] = true; });
  if (!have[D.TITLES.IMPROVE]) form.addParagraphTextItem().setTitle(D.TITLES.IMPROVE);

  // Remove deprecated items that older code versions added to the live form
  // (e.g. the "open loops" question was dropped as redundant with "担忧").
  const deprecated = ['今天有什么没做完、或一直记挂的事？'];
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    if (deprecated.indexOf(String(items[i].getTitle())) !== -1) {
      form.deleteItem(items[i]);
    }
  }

  // Idempotent trigger install: delete stale copies, recreate one.
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'handleDiarySubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('handleDiarySubmit').forForm(form).onFormSubmit().create();

  toast_('Diary set up — form + submit trigger ready.', 'Weekly Plan');
  return 'Diary form and submit trigger ready.';
}

/**
 * Scaffold the Diary sheet (header, freeze, column widths). Called from
 * setUpSheets. Idempotent — re-running only touches the header row.
 */
function setUpDiarySheet_() {
  const sheet = getOrCreateSheet_(CONFIG.DIARY_SHEET);
  const D = CONFIG.DIARY;
  sheet
    .getRange(D.COLS.DATE + '1:' + D.COLS.IMPROVE + '1')
    .setValues([['date', 'mood', 'worry', 'highlight', 'tomorrow_plan', 'submitted_at', 'updated_at', 'improve']])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);  // date
  sheet.setColumnWidth(3, 260);  // worry
  sheet.setColumnWidth(4, 260);  // highlight
  sheet.setColumnWidth(5, 260);  // tomorrow_plan
  sheet.setColumnWidth(8, 260);  // improve
}

// ----------------------------------------------------------------------
// Reminder
// ----------------------------------------------------------------------

/**
 * Build a form URL prefilled with today's date. Requires the forms scope.
 * @param {string} formId
 * @returns {string}
 */
function buildDiaryPrefilledUrl_(formId) {
  const form = FormApp.openById(formId);
  const fr = form.createResponse();
  const dateItems = form.getItems(FormApp.ItemType.DATE);
  if (dateItems.length) {
    // createResponse wants a Date (not a string) for a DateItem.
    fr.withItemResponse(dateItems[0].asDateItem().createResponse(new Date()));
  }
  return fr.toPrefilledUrl();
}

/**
 * Send the standalone diary reminder (Telegram + prefilled form link). Since
 * the diary became a morning plan this is NOT scheduled any more — the same
 * link rides along in the ~8:30 morning message (Telegram.js). Kept for the
 * "Send diary reminder now" menu item and for re-adding a separate trigger.
 * Deduped by day and enable-gated unless `force` (menu path).
 * @param {boolean} [force]
 * @returns {string} status message
 */
function sendDiaryReminder_(force) {
  if (!force && !CONFIG.DIARY.ENABLED) return 'Diary reminder disabled (Config.DIARY.ENABLED).';
  if (!telegramConfigured_()) return 'Telegram not configured — run "Set up Telegram…" first.';

  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty(CONFIG.DIARY.FORM_ID_PROP);
  if (!formId) return 'Diary form not set up — run "Set up Diary…" first.';

  const today = formatDateYYYYMMDD(new Date());
  if (!force && props.getProperty(CONFIG.DIARY.LAST_SENT_PROP) === today) {
    return 'Diary reminder already sent today.';
  }

  const url = buildDiaryPrefilledUrl_(formId);
  // Plain text (parseMode null) so the &-heavy prefilled URL is never
  // mangled by Telegram's HTML parser.
  sendTelegramMessage_('🌅 今日要点 — 花两分钟定下今天\n\n' + url, null);

  props.setProperty(CONFIG.DIARY.LAST_SENT_PROP, today);
  return 'Diary reminder sent.';
}

/**
 * Menu-callable manual send (forces past the dedup + enable flag).
 * @returns {string} status message
 */
function sendDiaryReminderFromUI() {
  let msg;
  try {
    msg = sendDiaryReminder_(true);
  } catch (e) {
    msg = 'Diary reminder failed: ' + (e && e.message ? e.message : e);
  }
  toast_(msg, 'Weekly Plan');
  return msg;
}

// ----------------------------------------------------------------------
// Form submit handler (installable onFormSubmit trigger)
// ----------------------------------------------------------------------

/**
 * onFormSubmit entry. Parses the responses (matched by item TITLE), upserts
 * the Diary sheet row, and refreshes the Weekly top band. The habit check and
 * the chronicler badge are intentionally deferred to spreadsheet-context runs.
 * @param {GoogleAppsScript.Events.FormOnSubmit} e
 */
function handleDiarySubmit(e) {
  try {
    const response = e && e.response;
    if (!response) return;
    const responses = response.getItemResponses() || [];

    // Date — prefer the prefilled Date item; fall back to the submit
    // timestamp with the late-night rule.
    let dateStr = null;
    const dateResp = findDiaryResponseByTitle_(responses, CONFIG.DIARY.TITLES.DATE);
    if (dateResp && dateResp.getResponse() instanceof Date) {
      dateStr = formatDateYYYYMMDD(dateResp.getResponse());
    }
    if (!dateStr) {
      let ts = null;
      try {
        ts = response.getTimestamp();
      } catch (err) {
        ts = null;
      }
      if (ts instanceof Date) {
        const d = new Date(ts);
        if (d.getHours() < CONFIG.DIARY.LATE_NIGHT_MAX_HOUR) d.setDate(d.getDate() - 1);
        dateStr = formatDateYYYYMMDD(d);
      }
    }
    if (!dateStr) return;

    const now = formatDateTime(new Date());
    const T = CONFIG.DIARY.TITLES;
    upsertDiaryRow_(
      dateStr,
      moodFromResponse_(responses),
      textFromResponse_(responses, T.KEY_POINTS),
      textFromResponse_(responses, T.HIGHLIGHT),
      textFromResponse_(responses, T.PLAN),
      textFromResponse_(responses, T.IMPROVE),
      now
    );

    refreshWeeklyFocusBand_();
    toast_('今日日记已记录', 'Weekly Plan');
  } catch (err) {
    Logger.log('handleDiarySubmit: ' + (err && err.message ? err.message : err));
  }
}

/**
 * Find an ItemResponse by its item's title (robust to form reordering).
 * @param {GoogleAppsScript.Forms.ItemResponse[]} responses
 * @param {string} title
 * @returns {GoogleAppsScript.Forms.ItemResponse|null}
 */
function findDiaryResponseByTitle_(responses, title) {
  for (let i = 0; i < responses.length; i++) {
    const item = responses[i].getItem();
    const t = item && item.getTitle ? item.getTitle() : '';
    if (String(t) === String(title)) return responses[i];
  }
  return null;
}

/**
 * Mood select → 1..5 (index+1), defaulting to 3 for missing/unknown values.
 * @param {GoogleAppsScript.Forms.ItemResponse[]} responses
 * @returns {number}
 */
function moodFromResponse_(responses) {
  const emoji = textFromResponse_(responses, CONFIG.DIARY.TITLES.MOOD);
  const idx = CONFIG.DIARY.MOOD_EMOJI.indexOf(emoji);
  return idx >= 0 ? idx + 1 : 3;
}

/**
 * Every title a field can arrive under: its current title plus any legacy
 * titles that CONFIG.DIARY.RENAMED_TITLES maps onto it. This is what lets the
 * code deploy BEFORE the live form is migrated by "Set up Diary…" — the old
 * item title still resolves, so a submission can't blank the column.
 * @param {string} title current title
 * @returns {string[]} current title first, then legacy titles
 */
function diaryTitleCandidates_(title) {
  const renames = CONFIG.DIARY.RENAMED_TITLES || {};
  const out = [title];
  for (const old in renames) {
    if (renames[old] === title) out.push(old);
  }
  return out;
}

/**
 * Free-text response → string ('' when absent). Matches the current title, then
 * its legacy titles (see diaryTitleCandidates_).
 * @param {GoogleAppsScript.Forms.ItemResponse[]} responses
 * @param {string} title
 * @returns {string}
 */
function textFromResponse_(responses, title) {
  const candidates = diaryTitleCandidates_(title);
  for (let i = 0; i < candidates.length; i++) {
    const resp = findDiaryResponseByTitle_(responses, candidates[i]);
    const v = resp && resp.getResponse();
    if (v !== undefined && v !== null && String(v) !== '') return String(v);
  }
  return '';
}

// ----------------------------------------------------------------------
// Diary sheet storage
// ----------------------------------------------------------------------

/**
 * Upsert one diary row by date (column A). New rows get submitted_at +
 * updated_at; existing rows keep submitted_at and touch only updated_at.
 * Runs under the document lock.
 * @param {string} dateStr YYYYMMDD
 * @param {number} mood 1–5
 * @param {string} worry
 * @param {string} highlight
 * @param {string} plan
 * @param {string} improve
 * @param {string} nowStr formatDateTime timestamp
 * @returns {{created:boolean, firstEver:boolean}}
 */
function upsertDiaryRow_(dateStr, mood, worry, highlight, plan, improve, nowStr) {
  const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
  if (!sheet) return { created: false, firstEver: false };
  const D = CONFIG.DIARY;
  const start = D.DATA_START_ROW;

  return withLock_(function () {
    const lastRow = getLastRowInColumn_(sheet, 1);
    const dates = lastRow >= start
      ? sheet.getRange(start, 1, lastRow - start + 1, 1).getValues()
      : [];
    let row = -1;
    for (let i = 0; i < dates.length; i++) {
      if (String(dates[i][0]) === String(dateStr)) {
        row = start + i;
        break;
      }
    }

    if (row >= 0) {
      // Existing: update mood..plan (B..E) + improve (H) + updated_at (G);
      // keep submitted_at (F) and the pre-existing fields.
      sheet.getRange(row, 2, 1, 4).setValues([[mood, worry, highlight, plan]]);
      sheet.getRange(D.COLS.IMPROVE + row).setValue(improve);
      sheet.getRange(D.COLS.UPDATED_AT + row).setValue(nowStr);
      return { created: false, firstEver: false };
    }

    const newRow = Math.max(lastRow + 1, start);
    const firstEver = lastRow < start;
    sheet.getRange(newRow, 1, 1, 8).setValues([[
      dateStr, mood, worry, highlight, plan, nowStr, nowStr, improve,
    ]]);
    return { created: true, firstEver: firstEver };
  });
}

/**
 * Diary row for a date, or null.
 * @param {string} dateStr YYYYMMDD
 * @returns {Object|null}
 */
function getDiaryForDate_(dateStr) {
  const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
  if (!sheet) return null;
  const D = CONFIG.DIARY;
  const lastRow = getLastRowInColumn_(sheet, 1);
  if (lastRow < D.DATA_START_ROW) return null;
  const dates = sheet.getRange(D.DATA_START_ROW, 1, lastRow - D.DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    if (String(dates[i][0]) === String(dateStr)) return readDiaryRow_(sheet, D.DATA_START_ROW + i);
  }
  return null;
}

/**
 * The diary row with the greatest date (rows may not be date-ordered), or null.
 * @returns {Object|null}
 */
function getMostRecentDiary_() {
  const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
  if (!sheet) return null;
  const D = CONFIG.DIARY;
  const lastRow = getLastRowInColumn_(sheet, 1);
  if (lastRow < D.DATA_START_ROW) return null;
  const dates = sheet.getRange(D.DATA_START_ROW, 1, lastRow - D.DATA_START_ROW + 1, 1).getValues();
  let best = null;
  for (let i = 0; i < dates.length; i++) {
    const ds = String(dates[i][0] || '');
    if (!ds) continue;
    if (!best || ds > best.date) best = { date: ds, row: D.DATA_START_ROW + i };
  }
  return best ? readDiaryRow_(sheet, best.row) : null;
}

/**
 * All diary rows (in sheet order).
 * @returns {Object[]}
 */
function getAllDiaries_() {
  const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
  if (!sheet) return [];
  const D = CONFIG.DIARY;
  const lastRow = getLastRowInColumn_(sheet, 1);
  const out = [];
  for (let row = D.DATA_START_ROW; row <= lastRow; row++) {
    const d = readDiaryRow_(sheet, row);
    if (d.date) out.push(d);
  }
  return out;
}

/**
 * Read a full diary row as an object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row 1-based
 * @returns {Object}
 */
function readDiaryRow_(sheet, row) {
  const vals = sheet.getRange(row, 1, 1, 8).getValues()[0];
  return {
    date: String(vals[0] || ''),
    mood: parseMood_(vals[1]),
    worry: String(vals[2] || ''),
    highlight: String(vals[3] || ''),
    tomorrow_plan: String(vals[4] || ''),
    submitted_at: String(vals[5] || ''),
    updated_at: String(vals[6] || ''),
    improve: String(vals[7] || ''),
  };
}

/** @param {*} v @returns {number|null} */
function parseMood_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? null : n;
}

// ----------------------------------------------------------------------
// Web App save (in-page editor on the local reader)
// ----------------------------------------------------------------------

/**
 * Save one diary row submitted from the local reader page (Web App POST
 * `action=diary`). Deliberately reuses upsertDiaryRow_ — the exact same write
 * primitive the Google Form trigger uses — so both entry points land on the
 * same Diary sheet row, keyed by date (D.DATA_START_ROW + column A).
 *
 * The "写日记" habit is marked here through processPendingDiaryHabits_ (the
 * spreadsheet-context helper the ~5am maintenance run also calls), so a save
 * from the page is scored exactly like a form submission — it just doesn't
 * wait for the next daily run. The helper is idempotent (per-date cell guard).
 *
 * @param {Object} p {date:'YYYY-MM-DD'|'YYYYMMDD', mood, worry, highlight, tomorrow_plan, improve}
 * @returns {{ok:boolean, created?:boolean, row?:Object, error?:string}}
 */
function saveDiaryFromWeb_(p) {
  const dateStr = diaryDateKeyFromInput_(p && p.date);
  if (!dateStr) {
    return { ok: false, error: 'invalid date (want YYYY-MM-DD): ' + ((p && p.date) || '') };
  }

  const res = upsertDiaryRow_(
    dateStr,
    diaryMoodFromInput_(p && p.mood),
    diaryTextFromInput_(p && p.worry),
    diaryTextFromInput_(p && p.highlight),
    diaryTextFromInput_(p && p.tomorrow_plan),
    diaryTextFromInput_(p && p.improve),
    formatDateTime(new Date())
  );

  refreshWeeklyFocusBand_();
  try {
    processPendingDiaryHabits_();
  } catch (e) {
    // best-effort: the daily maintenance run retries the same watermark path
    Logger.log('saveDiaryFromWeb_ habit: ' + (e && e.message ? e.message : e));
  }

  const row = getDiaryForDate_(dateStr);
  return { ok: true, created: res.created, row: row ? diaryRowToIso_(row) : null };
}

/**
 * Normalize an editor date input to the sheet's YYYYMMDD key, rejecting
 * impossible dates (e.g. 2026-02-31) and non-date junk. Accepts both the ISO
 * form the page sends and the raw sheet key.
 * @param {*} v
 * @returns {string|null} YYYYMMDD or null
 */
function diaryDateKeyFromInput_(v) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  let y, m, d;
  let mt = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (mt) {
    y = parseInt(mt[1], 10); m = parseInt(mt[2], 10); d = parseInt(mt[3], 10);
  } else {
    mt = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (!mt) return null;
    y = parseInt(mt[1], 10); m = parseInt(mt[2], 10); d = parseInt(mt[3], 10);
  }
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  // Round-trip through Date to reject overflow like 2026-02-31 (→ Mar 3).
  const dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
  return String(y) + String(m).padStart(2, '0') + String(d).padStart(2, '0');
}

/**
 * Editor mood input → 1..5, defaulting to 3 (same default as moodFromResponse_).
 * @param {*} v
 * @returns {number}
 */
function diaryMoodFromInput_(v) {
  const n = parseInt(v, 10);
  return (n >= 1 && n <= 5) ? n : 3;
}

/**
 * Editor free-text input → trimmed string. Capped well under Sheets' 50k cell
 * limit so a runaway paste can't blow up the row or the response payload.
 * @param {*} v
 * @returns {string}
 */
function diaryTextFromInput_(v) {
  if (v === undefined || v === null) return '';
  return String(v).trim().slice(0, 10000);
}

/**
 * readDiaryRow_ (YYYYMMDD key) → the ISO-dated row shape the local reader /
 * SQLite expect, so the page can show the saved row without a full reload.
 * @param {Object} row
 * @returns {Object}
 */
function diaryRowToIso_(row) {
  const ds = String(row.date || '');
  return {
    date: ds.slice(0, 4) + '-' + ds.slice(4, 6) + '-' + ds.slice(6, 8),
    mood: row.mood,
    worry: row.worry,
    highlight: row.highlight,
    tomorrow_plan: row.tomorrow_plan,
    improve: row.improve,
    submitted_at: row.submitted_at,
    updated_at: row.updated_at,
  };
}

// ----------------------------------------------------------------------
// Weekly top band
// ----------------------------------------------------------------------

/**
 * Write the most recent diary's 今日要点 (row 1) and 明日计划 (row 3) into the
 * Weekly sheet's frozen top band, so they're always visible. With the morning
 * diary the most recent entry is usually this morning's. Never inserts rows —
 * the grid row indices are hardcoded in Weekly.js.
 */
function refreshWeeklyFocusBand_() {
  const weekly = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!weekly) return;
  const diary = getMostRecentDiary_();
  writeBandCell_(
    weekly,
    CONFIG.DIARY.BAND_WORRY_RANGE,
    diary && diary.worry ? '📌 今日要点：' + diary.worry : '📌 还没有今日要点 — 早上记一条'
  );
  writeBandCell_(
    weekly,
    CONFIG.DIARY.BAND_PLAN_RANGE,
    diary && diary.tomorrow_plan ? '🎯 ' + diary.tomorrow_plan : '🎯 还没有明日计划'
  );
}

/**
 * Merge (if needed) and fill a single band cell; setValue on the top-left
 * cell (merged ranges reject setValues).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} weekly
 * @param {string} a1 e.g. 'C1:P1'
 * @param {string} text
 */
function writeBandCell_(weekly, a1, text) {
  try {
    const range = weekly.getRange(a1);
    if (!range.isPartOfMerge()) range.merge();
    range
      .getCell(1, 1)
      .setValue(text)
      .setWrap(true)
      .setVerticalAlignment('middle');
    weekly.setRowHeight(range.getRow(), CONFIG.DIARY.BAND_ROW_HEIGHT);
  } catch (e) {
    Logger.log('writeBandCell_ ' + a1 + ': ' + (e && e.message ? e.message : e));
  }
}

// ----------------------------------------------------------------------
// Deferred habit + badge (spreadsheet-context only)
// ----------------------------------------------------------------------

/**
 * Mark the "写日记" habit done for every diary date newer than the watermark
 * (and not already marked). Runs in a spreadsheet context (runDailyInitCore_ /
 * initialize) so XP/quest/boss side effects apply to the right document.
 * Idempotent: the per-date cell guard prevents re-counting even if the
 * watermark write lags.
 */
function processPendingDiaryHabits_() {
  const D = CONFIG.DIARY;
  const props = PropertiesService.getScriptProperties();
  const lastDate = props.getProperty(D.HABIT_LAST_DATE_PROP) || '';

  const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
  if (!sheet) return;
  const lastRow = getLastRowInColumn_(sheet, 1);
  if (lastRow < D.DATA_START_ROW) return;

  const today = formatDateYYYYMMDD(new Date());
  const dates = [];
  const vals = sheet.getRange(D.DATA_START_ROW, 1, lastRow - D.DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const ds = String(vals[i][0] || '');
    if (ds && ds > lastDate && ds <= today) dates.push(ds);
  }
  if (!dates.length) return;
  dates.sort();

  const habitRow = findHabitRowByName_(D.HABIT_NAME);
  if (habitRow < 0) return;

  for (let i = 0; i < dates.length; i++) {
    if (isHabitDoneForDate_(habitRow, dates[i])) continue;
    recordHabitDone(habitRow, diaryDateFromStr_(dates[i]));
  }

  // Advance the watermark (best-effort; the cell guard keeps this idempotent).
  tryWithLock_(function () {
    const cur = props.getProperty(D.HABIT_LAST_DATE_PROP) || '';
    props.setProperty(D.HABIT_LAST_DATE_PROP, dates[dates.length - 1] > cur ? dates[dates.length - 1] : cur);
  });
}

/**
 * True when the habit's cell for that date already has a completion (guard
 * against re-counting). Dates outside the 14-day window return true (skip).
 * @param {number} habitRow 1-based
 * @param {string} dateStr YYYYMMDD
 * @returns {boolean}
 */
function isHabitDoneForDate_(habitRow, dateStr) {
  const habits = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!habits) return true;
  const H = CONFIG.HABITS;
  const dayIdx = findHabitsDayIndexForDate_(habits, diaryDateFromStr_(dateStr));
  if (dayIdx < 0) return true;
  const col = columnLetterToIndex(H.COLUMNS.DAY_START) + 1 + dayIdx;
  const v = habits.getRange(habitRow, col).getValue();
  return !!(v && parseFloat(v) > 0);
}

/**
 * Award the "chronicler" badge once the first diary entry exists. Runs in a
 * spreadsheet context (runDailyInitCore_); awardBadge_ is idempotent.
 */
function maybeAwardChronicler_() {
  try {
    const sheet = getSheetByName_(CONFIG.DIARY_SHEET);
    if (!sheet) return;
    if (getLastRowInColumn_(sheet, 1) >= CONFIG.DIARY.DATA_START_ROW) {
      awardBadge_('chronicler');
    }
  } catch (e) {
    // best-effort
  }
}

/** @param {string} dateStr YYYYMMDD @returns {Date} local midnight */
function diaryDateFromStr_(dateStr) {
  return new Date(
    parseInt(dateStr.slice(0, 4), 10),
    parseInt(dateStr.slice(4, 6), 10) - 1,
    parseInt(dateStr.slice(6, 8), 10)
  );
}

// ----------------------------------------------------------------------
// WebApp payload
// ----------------------------------------------------------------------

/**
 * Diary rows for the WebApp ?view=diary endpoint: newest first, ISO dates,
 * and a per-date system snapshot joined from the Summary + Habits sheets.
 * @param {string} [since] YYYYMMDD inclusive
 * @param {string} [to] YYYYMMDD inclusive
 * @param {number} [limit] max rows (0 = unlimited)
 * @returns {Object[]}
 */
function buildDiaryViewPayload_(since, to, limit) {
  const all = getAllDiaries_();
  const out = [];
  for (let i = 0; i < all.length; i++) {
    const d = all[i];
    if (since && d.date < since) continue;
    if (to && d.date > to) continue;

    const iso = d.date.slice(0, 4) + '-' + d.date.slice(4, 6) + '-' + d.date.slice(6, 8);
    const y = parseInt(d.date.slice(0, 4), 10);
    const m = parseInt(d.date.slice(4, 6), 10) - 1;
    const day = parseInt(d.date.slice(6, 8), 10);

    let summary = null;
    try {
      summary = getSummaryForDate_(d.date);
    } catch (e) {
      summary = null;
    }
    let habitsDone = 0;
    try {
      habitsDone = habitsCompletedForDate_(new Date(y, m, day)) || 0;
    } catch (e) {
      habitsDone = 0;
    }

    out.push({
      date: iso,
      mood: d.mood,
      worry: d.worry,
      highlight: d.highlight,
      tomorrow_plan: d.tomorrow_plan,
      improve: d.improve,
      submitted_at: d.submitted_at,
      updated_at: d.updated_at,
      summary: summary, // {positive, negative, total} | null
      habits_done: habitsDone,
    });
  }
  out.sort(function (a, b) {
    return a.date < b.date ? 1 : a.date > b.date ? -1 : 0;
  });
  if (limit > 0 && out.length > limit) out.length = limit;
  return out;
}

// ----------------------------------------------------------------------
// Debug
// ----------------------------------------------------------------------

/**
 * Simulate a form submission without touching the real form — fabricated
 * ItemResponses drive handleDiarySubmit end-to-end. Menu/sidebar callable.
 * @returns {string} status message
 */
function testDiarySubmitFromUI() {
  const mk = function (title, val) {
    return {
      getItem: function () {
        return { getTitle: function () { return title; } };
      },
      getResponse: function () { return val; },
    };
  };
  const response = {
    getTimestamp: function () { return new Date(); },
    getItemResponses: function () {
      const T = CONFIG.DIARY.TITLES;
      return [
        mk(T.DATE, new Date()),
        mk(T.MOOD, '😐'),
        mk(T.KEY_POINTS, 'test key points'),
        mk(T.HIGHLIGHT, 'test highlight'),
        mk(T.PLAN, 'test plan'),
        mk(T.IMPROVE, 'test improve'),
      ];
    },
  };
  handleDiarySubmit({ response: response });
  toast_('Diary test submit ran.', 'Weekly Plan');
  return 'Diary test submit ran.';
}
