/**
 * Daily Diary feature (Google Sheets edition).
 *
 * Every evening (~22:30) a Telegram reminder links a prefilled Google Form.
 * Submitting it upserts one row in the Diary sheet (keyed by date), refreshes
 * the Weekly top band ("😟 worry" on row 1, "🎯 plan" on row 3), and marks the
 * "写日记" habit done — deferred to a spreadsheet-context run.
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

    // Fixed order + Chinese titles. handleDiarySubmit matches by TITLE (not
    // index), so reordering items later won't misparse responses.
    form.addDateItem().setTitle('日期');
    form.addMultipleChoiceItem().setTitle('心情').setChoiceValues(D.MOOD_EMOJI);
    form.addParagraphTextItem().setTitle('担忧');
    form.addParagraphTextItem().setTitle('亮点');
    form.addParagraphTextItem().setTitle('明日计划');
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
    .getRange(D.COLS.DATE + '1:' + D.COLS.UPDATED_AT + '1')
    .setValues([['date', 'mood', 'worry', 'highlight', 'tomorrow_plan', 'submitted_at', 'updated_at']])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);  // date
  sheet.setColumnWidth(3, 260);  // worry
  sheet.setColumnWidth(4, 260);  // highlight
  sheet.setColumnWidth(5, 260);  // tomorrow_plan
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
 * Send the daily diary reminder (Telegram + prefilled form link). Deduped by
 * day and enable-gated unless `force` (menu path).
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
  sendTelegramMessage_('🌙 每日日记 — 记录今天，放空大脑\n\n' + url, null);

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
    const dateResp = findDiaryResponseByTitle_(responses, '日期');
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
    upsertDiaryRow_(
      dateStr,
      moodFromResponse_(responses),
      textFromResponse_(responses, '担忧'),
      textFromResponse_(responses, '亮点'),
      textFromResponse_(responses, '明日计划'),
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
  const resp = findDiaryResponseByTitle_(responses, '心情');
  const emoji = resp && resp.getResponse() ? String(resp.getResponse()) : '';
  const idx = CONFIG.DIARY.MOOD_EMOJI.indexOf(emoji);
  return idx >= 0 ? idx + 1 : 3;
}

/**
 * Free-text response → string ('' when absent).
 * @param {GoogleAppsScript.Forms.ItemResponse[]} responses
 * @param {string} title
 * @returns {string}
 */
function textFromResponse_(responses, title) {
  const resp = findDiaryResponseByTitle_(responses, title);
  const v = resp && resp.getResponse();
  return v === undefined || v === null ? '' : String(v);
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
 * @param {string} nowStr formatDateTime timestamp
 * @returns {{created:boolean, firstEver:boolean}}
 */
function upsertDiaryRow_(dateStr, mood, worry, highlight, plan, nowStr) {
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
      // Existing: update mood..plan (B..E) + updated_at (G), keep submitted_at (F).
      sheet.getRange(row, 2, 1, 4).setValues([[mood, worry, highlight, plan]]);
      sheet.getRange(D.COLS.UPDATED_AT + row).setValue(nowStr);
      return { created: false, firstEver: false };
    }

    const newRow = Math.max(lastRow + 1, start);
    const firstEver = lastRow < start;
    sheet.getRange(newRow, 1, 1, 7).setValues([[
      dateStr, mood, worry, highlight, plan, nowStr, nowStr,
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
  const vals = sheet.getRange(row, 1, 1, 7).getValues()[0];
  return {
    date: String(vals[0] || ''),
    mood: parseMood_(vals[1]),
    worry: String(vals[2] || ''),
    highlight: String(vals[3] || ''),
    tomorrow_plan: String(vals[4] || ''),
    submitted_at: String(vals[5] || ''),
    updated_at: String(vals[6] || ''),
  };
}

/** @param {*} v @returns {number|null} */
function parseMood_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? null : n;
}

// ----------------------------------------------------------------------
// Weekly top band
// ----------------------------------------------------------------------

/**
 * Write the most recent diary's worry (row 1) and tomorrow plan (row 3) into
 * the Weekly sheet's frozen top band, so they're always visible. Never
 * inserts rows — the grid row indices are hardcoded in Weekly.js.
 */
function refreshWeeklyFocusBand_() {
  const weekly = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!weekly) return;
  const diary = getMostRecentDiary_();
  writeBandCell_(
    weekly,
    CONFIG.DIARY.BAND_WORRY_RANGE,
    diary && diary.worry ? '😟 ' + diary.worry : '😟 还没有日记 — 今晚填一份'
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
      return [
        mk('日期', new Date()),
        mk('心情', '😐'),
        mk('担忧', 'test worry'),
        mk('亮点', 'test highlight'),
        mk('明日计划', 'test plan'),
      ];
    },
  };
  handleDiarySubmit({ response: response });
  toast_('Diary test submit ran.', 'Weekly Plan');
  return 'Diary test submit ran.';
}
