/**
 * Morning Telegram notification (Google Sheets edition).
 *
 * Sends one message each morning (~CONFIG.TELEGRAM.SEND_HOUR) with
 * yesterday's performance and today's Daily Quest (featured habit +
 * task), plus the combo / XP / weekly-boss lines — the same recap the
 * email build produces, formatted for Telegram's HTML parse mode.
 *
 * Driven by the time-based `morningTelegram` trigger (see Triggers.js),
 * NOT by sidebar init — so it fires once a day. A DocumentProperties
 * stamp (CONFIG.TELEGRAM.LAST_SENT_PROP) guarantees no duplicate sends.
 *
 * Bot credentials (token + chat id) are read from Script Properties and
 * are never stored in source. Set them once via setUpTelegram() (the
 * "Set up Telegram…" menu item) or the Script Properties editor.
 *
 * Requires the script.external_request OAuth scope (see appsscript.json).
 */

/**
 * Read the bot token + chat id from Script Properties.
 * @returns {{token:string, chatId:string}}
 */
function telegramCredentials_() {
  const props = PropertiesService.getScriptProperties();
  return {
    token: (props.getProperty(CONFIG.TELEGRAM.BOT_TOKEN_PROP) || '').trim(),
    chatId: (props.getProperty(CONFIG.TELEGRAM.CHAT_ID_PROP) || '').trim(),
  };
}

/**
 * True when both the bot token and chat id are present.
 * @returns {boolean}
 */
function telegramConfigured_() {
  const c = telegramCredentials_();
  return !!(c.token && c.chatId);
}

/**
 * Persist the bot token + chat id to Script Properties.
 * @param {string} token from @BotFather
 * @param {string} chatId numeric Telegram chat id
 */
function setTelegramCredentials_(token, chatId) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.TELEGRAM.BOT_TOKEN_PROP, String(token));
  props.setProperty(CONFIG.TELEGRAM.CHAT_ID_PROP, String(chatId));
}

/**
 * Escape the characters Telegram's HTML parse mode treats as markup, so
 * arbitrary habit/task/boss names render literally.
 * @param {string} s
 * @returns {string}
 */
function escapeHtmlTelegram_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Truncate a diary free-text field for the recap (~60 chars), with a
 * placeholder for empty values.
 * @param {string} s
 * @returns {string}
 */
function truncateDiaryText_(s) {
  const t = String(s || '').trim();
  if (!t) return '—';
  const MAX = 60;
  return t.length > MAX ? t.slice(0, MAX) + '…' : t;
}

/**
 * POST a message to the Telegram Bot API (sendMessage).
 * @param {string} text message body
 * @param {string|null} [parseMode] 'HTML' (default) for formatted text,
 *   or null/'' to send as plain text (use for untrusted/LLM content that
 *   may contain characters that would break HTML parsing).
 * @returns {string} the API response body
 * @throws {Error} when credentials are missing or the API returns non-2xx
 */
function sendTelegramMessage_(text, parseMode) {
  const creds = telegramCredentials_();
  if (!creds.token || !creds.chatId) {
    throw new Error('Telegram credentials not set — run "Set up Telegram…" first.');
  }
  const mode = parseMode === undefined ? 'HTML' : parseMode;
  const url = 'https://api.telegram.org/bot' + creds.token + '/sendMessage';
  const payload = {
    chat_id: creds.chatId,
    text: text,
    disable_web_page_preview: true,
  };
  if (mode) payload.parse_mode = mode;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Telegram API error ' + code + ': ' + body);
  }
  return body;
}

/**
 * Build the morning message (Telegram HTML). Since the diary became a morning
 * plan (2026-10-15) this single ~8:30 message carries both the recap and the
 * "今日要点" prompt, so there is one morning notification instead of two.
 *
 * The diary is written in the morning, so the previous day's row is what is
 * actionable now: its `明日计划` was written FOR today, and its 今日要点 is
 * yesterday's focus. Links go at the bottom (thumb-reachable).
 *
 * @returns {string}
 */
function buildMorningTelegramMessage_() {
  const tz = getSpreadsheet_().getSpreadsheetTimeZone();
  const now = new Date();
  const yDate = new Date(now);
  yDate.setDate(now.getDate() - 1);

  const quest = ensureDailyQuest_();
  const ys = getSummaryForDate_(formatDateYYYYMMDD(yDate));
  const habitsDone = habitsCompletedForDate_(yDate);

  const fmt = (n) => (Math.round(n * 100) / 100).toString();
  const esc = escapeHtmlTelegram_;
  const yLabel = Utilities.formatDate(yDate, tz, 'EEEE, MMM d');
  const tLabel = Utilities.formatDate(now, tz, 'EEEE, MMM d');
  const bonus = CONFIG.QUEST.BONUS_MULTIPLIER;

  const lines = [];
  lines.push('🌅 <b>Weekly Plan — Morning Plan</b>');
  lines.push('<i>' + esc(tLabel) + '</i>');
  lines.push('');

  // ---- Yesterday section ----
  lines.push('<b>Yesterday — ' + esc(yLabel) + '</b>');
  if (ys) {
    lines.push(
      'Total <b>' + fmt(ys.total) + '</b>  (➕ ' + fmt(ys.positive) + ' / ➖ ' + fmt(ys.negative) + ')'
    );
    lines.push('Habits completed: <b>' + habitsDone + '</b>');
  } else {
    lines.push('No activity logged. Fresh start today! 💪');
  }
  lines.push('');

  // ---- Yesterday's diary: its 明日计划 was written FOR today ----
  const diary = getDiaryForDate_(formatDateYYYYMMDD(yDate));
  if (diary) {
    lines.push('🎯 昨天为今天定的计划: <b>' + esc(truncateDiaryText_(diary.tomorrow_plan)) + '</b>');
    lines.push('📌 昨天写的今日要点: ' + esc(truncateDiaryText_(diary.worry)));
  } else {
    lines.push('<i>昨天没有日记 — 今天开始记第一条吧</i>');
  }
  lines.push('');

  // ---- Today's quest section ----
  const qHabit = (quest && quest.habitName) || '—';
  const qTask = (quest && quest.taskName) || '—';
  lines.push("⭐ <b>Today's Quest</b>");
  lines.push('🎯 Habit: <b>' + esc(qHabit) + '</b>');
  lines.push('📋 Task: <b>' + esc(qTask) + '</b>');
  lines.push('Complete both for a ×' + bonus + ' bonus!');

  // ---- Combo at-risk nudge ----
  const combo = getComboState_();
  if (combo.atRisk && combo.days > 0) {
    lines.push('');
    lines.push(
      '🔥 Your ' + combo.days + '-day quest combo is at risk — finish today to reach ×' +
        combo.multiplier.toFixed(1) + '!'
    );
  }

  // ---- XP / level line ----
  const xp = getXpStateFromUI();
  lines.push('');
  let xpLine =
    '🏆 Level ' + xp.level + ' — ' + xp.intoLevel + ' / ' + xp.levelSpan + ' XP to next';
  if (xp.badges && xp.badges.length) {
    xpLine += '  ·  ' + xp.badges.map((b) => b.emoji).join(' ');
  }
  lines.push(xpLine);

  // ---- Boss line ----
  const boss = getBossStateFromUI();
  if (boss) {
    const unit = boss.type === 'quests' ? ' quest(s)' : ' pts';
    if (boss.defeated) {
      lines.push(boss.emoji + ' Weekly Boss defeated: ' + esc(boss.name) + '! 🎉');
    } else {
      lines.push(
        boss.emoji + ' Weekly Boss: ' + esc(boss.name) + ' — ' + boss.remaining + unit +
          ' to go (' + boss.progress + '/' + boss.target + ')'
      );
    }
  }

  // ---- Action links (plain text: the prefill URL is &-heavy and would be
  //      mangled by HTML mode; and the sheet id is harmless to expose). ----
  const formUrl = diaryFormUrl_();
  const sheetUrl = dailySheetUrl_();
  lines.push('');
  if (formUrl) {
    lines.push('✍️ <b>写今日要点</b>（已预填今天）');
    lines.push(formUrl);
  }
  if (sheetUrl) {
    lines.push('📊 <b>打开 Weekly Plan</b>');
    lines.push(sheetUrl);
  }

  return lines.join('\n');
}

/**
 * Prefilled diary form URL for today, or '' when the form isn't set up.
 * @returns {string}
 */
function diaryFormUrl_() {
  const formId = PropertiesService.getScriptProperties().getProperty(CONFIG.DIARY.FORM_ID_PROP);
  if (!formId) return '';
  try {
    return buildDiaryPrefilledUrl_(formId);
  } catch (e) {
    Logger.log('diaryFormUrl_: ' + (e && e.message ? e.message : e));
    return '';
  }
}

/**
 * The bound spreadsheet's URL — the "open Weekly Plan" link for the morning
 * message. getSpreadsheet_() carries the openById fallback, so this also works
 * in a time-driven (no active spreadsheet) context. Falls back to building the
 * URL from the id if getUrl() isn't available under the current OAuth scopes.
 * @returns {string}
 */
function dailySheetUrl_() {
  try {
    const ss = getSpreadsheet_();
    try {
      const url = ss.getUrl();
      if (url) return url;
    } catch (e) {
      // fall through to the constructed URL
    }
    return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit';
  } catch (e) {
    Logger.log('dailySheetUrl_: ' + (e && e.message ? e.message : e));
    return '';
  }
}

/**
 * Send the morning Telegram recap. Respects the enable flag and the
 * once-per-day stamp unless `force` is true (used by the manual menu).
 * @param {boolean} [force] bypass the enabled flag + dedup stamp
 * @returns {string} status message
 */
function sendMorningTelegram_(force) {
  if (!force && !CONFIG.TELEGRAM.ENABLED) {
    return 'Telegram recap disabled (Config.TELEGRAM.ENABLED).';
  }
  if (!telegramConfigured_()) {
    return 'Telegram not configured — run "Set up Telegram…" first.';
  }

  const props = PropertiesService.getDocumentProperties();
  const today = formatDateYYYYMMDD(new Date());
  if (!force && props.getProperty(CONFIG.TELEGRAM.LAST_SENT_PROP) === today) {
    return 'Telegram recap already sent today.';
  }

  sendTelegramMessage_(buildMorningTelegramMessage_());

  props.setProperty(CONFIG.TELEGRAM.LAST_SENT_PROP, today);
  return 'Telegram morning recap sent.';
}

/**
 * Menu-callable credential setup: prompt for the bot token + chat id,
 * store them in Script Properties, and send a test message so the user
 * can confirm the connection immediately.
 */
function setUpTelegram() {
  const ui = SpreadsheetApp.getUi();

  const tokenResp = ui.prompt(
    'Set up Telegram (1/2)',
    'Paste your bot token from @BotFather:',
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenResp.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenResp.getResponseText().trim();
  if (!token) {
    ui.alert('Bot token is required.');
    return;
  }

  const chatResp = ui.prompt(
    'Set up Telegram (2/2)',
    'Paste your numeric chat id (message @userinfobot to find it):',
    ui.ButtonSet.OK_CANCEL
  );
  if (chatResp.getSelectedButton() !== ui.Button.OK) return;
  const chatId = chatResp.getResponseText().trim();
  if (!chatId) {
    ui.alert('Chat id is required.');
    return;
  }

  setTelegramCredentials_(token, chatId);

  try {
    sendTelegramMessage_(
      "✅ Weekly Plan is connected. You'll get your morning recap around " +
        CONFIG.TELEGRAM.SEND_HOUR + ':00 each day.'
    );
    ui.alert('Telegram connected! A test message was sent — check your Telegram app.');
  } catch (e) {
    ui.alert(
      'Credentials saved, but the test message failed:\n\n' +
        (e && e.message ? e.message : e) +
        '\n\nDouble-check the token and chat id, then try "Send Telegram recap now".'
    );
  }
}

/**
 * Menu/sidebar-callable manual send (forces past the dedup + enable
 * flag) so the user can preview the recap any time.
 * @returns {string} status message
 */
function sendTelegramNowFromUI() {
  let msg;
  try {
    msg = sendMorningTelegram_(true);
  } catch (e) {
    msg = 'Telegram send failed: ' + (e && e.message ? e.message : e);
  }
  toast_(msg, 'Weekly Plan');
  return msg;
}
