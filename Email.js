/**
 * Morning summary email (Google Sheets edition).
 *
 * Sends one email each morning with yesterday's performance and today's
 * Daily Quest. Driven by the daily time-based trigger (dailyMaintenance,
 * ~5am) — NOT by sidebar init — so it fires once a day. A DocumentProperties
 * stamp (CONFIG.EMAIL.LAST_SENT_PROP) guarantees no duplicate sends per day.
 *
 * Requires the script.send_mail OAuth scope (see appsscript.json).
 */

/**
 * Resolve the recipient address: the configured one, else the script
 * owner's account.
 * @returns {string}
 */
function emailRecipient_() {
  const configured = (CONFIG.EMAIL.RECIPIENT || '').trim();
  if (configured) return configured;
  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Count habit completions on a given date (within the 14-day window).
 * @param {Date} date
 * @returns {number}
 */
function habitsCompletedForDate_(date) {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return 0;
  const H = CONFIG.HABITS;
  const dayIndex = findHabitsDayIndexForDate_(sheet, date);
  if (dayIndex < 0) return 0;
  const start = H.DATA_START_ROW;
  const last = getLastHabitRow_();
  if (last < start) return 0;
  const col = columnLetterToIndex(H.COLUMNS.DAY_START) + 1 + dayIndex; // 1-based
  const vals = sheet.getRange(start, col, last - start + 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < vals.length; i++) {
    const n = parseInt(vals[i][0], 10);
    if (n && n > 0) count++;
  }
  return count;
}

/**
 * Build the morning email subject + bodies. Always returns content (the
 * quest section is present even with no yesterday data).
 * @returns {{subject:string, htmlBody:string, textBody:string}}
 */
function buildMorningEmail_() {
  const tz = getSpreadsheet_().getSpreadsheetTimeZone();
  const now = new Date();
  const yDate = new Date(now);
  yDate.setDate(now.getDate() - 1);

  const quest = ensureDailyQuest_();
  const ys = getSummaryForDate_(formatDateYYYYMMDD(yDate));
  const habitsDone = habitsCompletedForDate_(yDate);

  const fmt = (n) => (Math.round(n * 100) / 100).toString();
  const yLabel = Utilities.formatDate(yDate, tz, 'EEEE, MMM d');
  const tLabel = Utilities.formatDate(now, tz, 'EEEE, MMM d');
  const bonus = CONFIG.QUEST.BONUS_MULTIPLIER;

  // ---- Yesterday section ----
  let ydHtml;
  let ydText;
  if (ys) {
    ydHtml =
      '<table style="border-collapse:collapse;width:100%;max-width:360px">' +
      '<tr>' +
      '<td style="padding:8px;background:#e8f5e9;border-radius:8px;text-align:center">' +
      '<div style="font-size:11px;color:#2e7d32;text-transform:uppercase">Positive</div>' +
      '<div style="font-size:20px;font-weight:700;color:#2e7d32">' + fmt(ys.positive) + '</div></td>' +
      '<td style="width:8px"></td>' +
      '<td style="padding:8px;background:#ede7f6;border-radius:8px;text-align:center">' +
      '<div style="font-size:11px;color:#5e35b1;text-transform:uppercase">Total</div>' +
      '<div style="font-size:20px;font-weight:700;color:#5e35b1">' + fmt(ys.total) + '</div></td>' +
      '<td style="width:8px"></td>' +
      '<td style="padding:8px;background:#ffebee;border-radius:8px;text-align:center">' +
      '<div style="font-size:11px;color:#c62828;text-transform:uppercase">Negative</div>' +
      '<div style="font-size:20px;font-weight:700;color:#c62828">' + fmt(ys.negative) + '</div></td>' +
      '</tr></table>' +
      '<p style="margin:8px 0 0;color:#555">Habits completed: <strong>' + habitsDone + '</strong></p>';
    ydText =
      'Yesterday (' + yLabel + '): total ' + fmt(ys.total) +
      ' (+' + fmt(ys.positive) + ' / ' + fmt(ys.negative) + '), habits completed ' + habitsDone + '.';
  } else {
    ydHtml = '<p style="color:#888">No activity was logged yesterday. Fresh start today! 💪</p>';
    ydText = 'Yesterday (' + yLabel + '): no activity logged.';
  }

  // ---- Today's quest section ----
  const qHabit = (quest && quest.habitName) || '—';
  const qTask = (quest && quest.taskName) || '—';
  const questHtml =
    '<div style="background:#fff8e1;border-radius:10px;padding:14px;margin-top:8px">' +
    '<div style="font-size:14px;font-weight:700;color:#a87900;margin-bottom:8px">⭐ Today\'s Quest</div>' +
    '<div style="margin:4px 0"><span style="font-size:11px;color:#a87900;text-transform:uppercase">Habit</span><br>' +
    '<span style="font-size:15px;font-weight:600;color:#333">🎯 ' + escapeHtmlEmail_(qHabit) + '</span></div>' +
    '<div style="margin:8px 0 4px"><span style="font-size:11px;color:#a87900;text-transform:uppercase">Task</span><br>' +
    '<span style="font-size:15px;font-weight:600;color:#333">📋 ' + escapeHtmlEmail_(qTask) + '</span></div>' +
    '<div style="font-size:12px;color:#a87900;margin-top:8px">Complete both for a ×' + bonus + ' bonus!</div>' +
    '</div>';
  const questText =
    "Today's Quest (" + tLabel + '): Habit "' + qHabit + '", Task "' + qTask +
    '" — complete both for a x' + bonus + ' bonus.';

  // ---- XP / level line ----
  const xp = getXpStateFromUI();
  const xpHtml =
    '<p style="margin:10px 0 0;color:#5e35b1;font-weight:600">🏆 Level ' + xp.level +
    ' — ' + xp.intoLevel + ' / ' + xp.levelSpan + ' XP to next' +
    (xp.badges.length ? '  ·  ' + xp.badges.map((b) => b.emoji).join(' ') : '') + '</p>';
  const xpText = 'Level ' + xp.level + ' (' + xp.intoLevel + '/' + xp.levelSpan + ' XP to next).';

  // ---- Combo at-risk nudge ----
  const combo = getComboState_();
  let comboHtml = '';
  let comboText = '';
  if (combo.atRisk && combo.days > 0) {
    comboHtml =
      '<p style="margin:8px 0 0;color:#c62828;font-weight:600">🔥 Your ' + combo.days +
      '-day quest combo is at risk — finish today\'s quest to reach ×' + combo.multiplier.toFixed(1) + '!</p>';
    comboText = 'Combo: ' + combo.days + '-day streak at risk — finish today to keep it.';
  }

  // ---- Boss line ----
  const boss = getBossStateFromUI();
  let bossHtml = '';
  let bossText = '';
  if (boss) {
    const unit = boss.type === 'quests' ? ' quest(s)' : ' pts';
    if (boss.defeated) {
      bossHtml = '<p style="margin:8px 0 0;color:#2e7d32;font-weight:600">' + boss.emoji +
        ' Weekly Boss defeated: ' + boss.name + '! 🎉</p>';
      bossText = 'Weekly Boss: ' + boss.name + ' defeated!';
    } else {
      bossHtml = '<p style="margin:8px 0 0;color:#b3375f;font-weight:600">' + boss.emoji +
        ' Weekly Boss: ' + boss.name + ' — ' + boss.remaining + unit + ' to go (' +
        boss.progress + '/' + boss.target + ')</p>';
      bossText = 'Weekly Boss: ' + boss.name + ' — ' + boss.remaining + unit + ' to go (' +
        boss.progress + '/' + boss.target + ').';
    }
  }

  const htmlBody =
    '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:480px;margin:0 auto;color:#333">' +
    '<h2 style="color:#5e35b1;margin:0 0 4px">📅 Weekly Plan — Morning Recap</h2>' +
    '<p style="color:#888;margin:0 0 16px">' + tLabel + '</p>' +
    '<h3 style="margin:0 0 8px;color:#333">Yesterday — ' + yLabel + '</h3>' +
    ydHtml +
    questHtml +
    comboHtml +
    xpHtml +
    bossHtml +
    '<p style="font-size:11px;color:#aaa;margin-top:20px">Sent by your Weekly Plan sheet. ' +
    'Turn this off in Config.js (EMAIL.ENABLED).</p>' +
    '</div>';

  const textBody =
    'Weekly Plan — Morning Recap (' + tLabel + ')\n\n' + ydText + '\n\n' + questText +
    (comboText ? '\n\n' + comboText : '') + '\n\n' + xpText +
    (bossText ? '\n\n' + bossText : '') + '\n';

  return {
    subject: 'Weekly Plan — ' + tLabel + ' recap & quest',
    htmlBody: htmlBody,
    textBody: textBody,
  };
}

/**
 * Minimal HTML escaper for untrusted habit/task names in the email body.
 * @param {string} s
 * @returns {string}
 */
function escapeHtmlEmail_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Send the morning summary email. Respects the enable flag and the
 * once-per-day stamp unless `force` is true (used by the manual menu item).
 * @param {boolean} [force] bypass the enabled flag + dedup stamp
 * @returns {string} status message
 */
function sendMorningEmail_(force) {
  if (!force && !CONFIG.EMAIL.ENABLED) return 'Morning email disabled (Config.EMAIL.ENABLED).';

  const props = PropertiesService.getDocumentProperties();
  const today = formatDateYYYYMMDD(new Date());
  if (!force && props.getProperty(CONFIG.EMAIL.LAST_SENT_PROP) === today) {
    return 'Morning email already sent today.';
  }

  const to = emailRecipient_();
  if (!to) return 'No email recipient resolved.';

  const mail = buildMorningEmail_();
  MailApp.sendEmail({
    to: to,
    subject: mail.subject,
    htmlBody: mail.htmlBody,
    body: mail.textBody,
  });

  props.setProperty(CONFIG.EMAIL.LAST_SENT_PROP, today);
  return 'Morning email sent to ' + to + '.';
}

/**
 * Menu/sidebar-callable manual send (forces past the dedup + enable flag),
 * so users can preview the email any time.
 * @returns {string} status message
 */
function sendSummaryEmailNowFromUI() {
  const msg = sendMorningEmail_(true);
  toast_(msg, 'Weekly Plan');
  return msg;
}
