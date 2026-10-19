/**
 * Web App 入口 — 部署为 GAS Web App（执行身份：我，访问权限：任何人）。
 *
 * GET  ?k=<auth> → 返回每日快照 JSON（供 Mac 定时拉取）
 * POST ?k=<auth> + JSON body → 执行操作（习惯打卡、日记保存等）
 *
 * POST action=diary 是日记网站（local/DiaryReader.html）内联编辑器的写入口：
 * 它走 saveDiaryFromWeb_ → upsertDiaryRow_，和 Telegram 推送的 Google 表单
 * 提交走的是同一个写路径，因此两种入口改的是同一张 Diary 表的同一行（按日期）。
 *
 * 要求 Utils.js 的 getSpreadsheet_() 已加入 openById fallback：
 *   PropertiesService.getScriptProperties().setProperty('spreadsheetId', '<id>')
 */

// ---- 工具函数 ----

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function unauth_() {
  return json_({ ok: false, error: 'unauthorized' });
}

function auth_(e) {
  var key = (e && e.parameter && e.parameter.k) || '';
  var expected = (PropertiesService.getScriptProperties().getProperty('syncAuthKey') || '');
  return !expected || key === expected;
}

/**
 * 严格鉴权：syncAuthKey 未设置时直接拒绝。用于承载私人内容的入口
 * （日记读取 ?view=diary、日记保存 action=diary）——宽松的 auth_() 会在
 * key 未配置时放行，不能用于私人内容。
 */
function strictAuth_(e) {
  var expected = (PropertiesService.getScriptProperties().getProperty(CONFIG.SYNC.AUTH_KEY_PROP) || '');
  var key = (e && e.parameter && e.parameter.k) || '';
  return !!expected && key === expected;
}

// ---- GET: 每日快照 ----

function doGet(e) {
  // Diary view: private content — strict auth (refuse when syncAuthKey unset),
  // enforced inside handleDiaryView_. Everything else uses the lenient auth_.
  if (e && e.parameter && e.parameter.view === 'diary') {
    return handleDiaryView_(e);
  }
  if (e && e.parameter && e.parameter.view === 'links') {
    return handleLinksView_(e);
  }
  if (e && e.parameter && e.parameter.view === 'schedule') {
    return handleScheduleView_(e);
  }
  if (!auth_(e)) return unauth_();
  try {
    var snapshot = buildSnapshot_();
    snapshot.ok = true;
    return json_(snapshot);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * ?view=links&k=KEY — the URLs the local reader needs (the Weekly Plan sheet
 * link on the diary page). Strict auth, same as the diary view: it identifies
 * the user's document. The local server caches it in local/.env, so a page
 * load never waits on Apps Script.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function handleLinksView_(e) {
  if (!strictAuth_(e)) return unauth_();
  try {
    return json_({ ok: true, weekly_plan_url: dailySheetUrl_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * ?view=schedule&k=KEY — today's timetable for the diary page's "今日排程"
 * panel. readTodayGrid_ reads the Weekly grid's column for today, and
 * Calendar.js writes synced meetings into those same task cells, so planned
 * blocks and calendar events come from one source.
 *
 * Deliberately light (no quest/boss/xp reads): the diary page fetches it on
 * every load, and the local server caches the response for a few minutes.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function handleScheduleView_(e) {
  if (!strictAuth_(e)) return unauth_();
  try {
    var tz = getSpreadsheet_().getSpreadsheetTimeZone();
    var now = new Date();
    var grid = readTodayGrid_();
    return json_({
      ok: true,
      day: Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      weekday: Utilities.formatDate(now, tz, 'EEEE'),
      slots: grid.slots,             // [{ time, task, score }]
      empty_slots: grid.empty_slots  // ['09:00', ...] unplanned 30-min slots
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * ?view=diary&k=KEY&since=YYYYMMDD&to=YYYYMMDD&limit=N[&callback=cb]
 * Diary rows (newest first) with a per-date system snapshot joined from the
 * Summary + Habits sheets. Strict auth: refuses when syncAuthKey is unset.
 * JSONP via `callback` (only needed if a browser page talks to the Web App
 * directly; the Mac bot fetches server-side and needs no CORS workaround).
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function handleDiaryView_(e) {
  if (!strictAuth_(e)) return unauth_();

  try {
    var p = (e && e.parameter) || {};
    var rows = buildDiaryViewPayload_(
      p.since || '',
      p.to || '',
      parseInt(p.limit, 10) || 0
    );
    var payload = { ok: true, diary: rows };

    var cb = p.callback || '';
    if (cb) {
      if (!/^[A-Za-z_$][\w$]*$/.test(cb)) {
        return json_({ ok: false, error: 'invalid callback' });
      }
      return ContentService
        .createTextOutput(cb + '(' + JSON.stringify(payload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json_(payload);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ---- POST: 操作入口 ----

function doPost(e) {
  var params;
  try {
    params = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad JSON body' });
  }

  try {
    // Diary saves carry private content → strict auth (refuse when the key is
    // unset), the same rule as the ?view=diary read.
    if (params.action === 'diary') {
      if (!strictAuth_(e)) return unauth_();
      return handleDiarySave_(params);
    }

    if (!auth_(e)) return unauth_();
    // { action: "habit", name: "健身" }
    if (params.action === 'habit' && params.name) {
      return handleHabitCheckin_(params.name);
    }
    return json_({ ok: false, error: 'unknown action: ' + (params.action || '') });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ---- 日记保存（网站内联编辑器） ----

/**
 * POST action=diary — 日记网站里点“保存”走这里。写入委托给
 * saveDiaryFromWeb_（→ upsertDiaryRow_），与 Google 表单提交共用同一条写
 * 路径，所以同一天无论是表单还是网站改的，都是 Diary 表的同一行。
 * Body: { action:'diary', date:'YYYY-MM-DD', mood:1-5, worry, highlight,
 *         tomorrow_plan, improve }
 * @param {Object} params
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function handleDiarySave_(params) {
  var res = saveDiaryFromWeb_(params || {});
  if (!res.ok) return json_({ ok: false, error: res.error });
  return json_({ ok: true, action: 'diary', created: res.created, row: res.row });
}

// ---- 习惯打卡 ----

function handleHabitCheckin_(name) {
  var sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return json_({ ok: false, error: 'Habits sheet not found' });

  var lastRow = getLastHabitRow_();
  var nameCol = columnLetterToIndex(CONFIG.HABITS.COLUMNS.HABIT_NAME) + 1; // B
  var names = sheet.getRange(CONFIG.HABITS.DATA_START_ROW, nameCol, lastRow - CONFIG.HABITS.DATA_START_ROW + 1, 1)
    .getValues().flat();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i]).trim() === name) {
      var row = CONFIG.HABITS.DATA_START_ROW + i;
      recordHabitDone(row);
      return json_({ ok: true, habit: name });
    }
  }
  return json_({ ok: false, error: 'habit not found: ' + name });
}

// ---- 快照构建 ----

function buildSnapshot_() {
  var tz = getSpreadsheet_().getSpreadsheetTimeZone();
  var today = new Date();
  var todayISO = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  var todayYYYYMMDD = formatDateYYYYMMDD(today);
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  var yesterdayISO = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');
  var yesterdayYYYYMMDD = formatDateYYYYMMDD(yesterday);

  // 确保 quest / boss 已初始化（幂等，sheet 打开时 5am trigger 通常已跑过）
  ensureDailyQuest_();
  ensureWeeklyBoss_();

  // ---- Summary ----
  var summary = {};
  var ys = getSummaryForDate_(yesterdayYYYYMMDD);
  if (ys) {
    summary[yesterdayISO] = {
      positive: ys.positive, negative: ys.negative, total: ys.total,
      habits_done: habitsCompletedForDate_(yesterday)
    };
  }
  var ts = getSummaryForDate_(todayYYYYMMDD);
  if (ts) {
    summary[todayISO] = {
      positive: ts.positive, negative: ts.negative, total: ts.total,
      habits_done: habitsCompletedForDate_(today)
    };
  }

  // ---- Quest ----
  var quest = getDailyQuestFromUI();

  // ---- XP / Badges / Boss ----
  var xp = getXpStateFromUI();
  var boss = getBossStateFromUI();

  // ---- Habits ----
  var habits = readHabits_(today);
  var habitsDoneCount = 0;
  for (var i = 0; i < habits.length; i++) {
    if (habits[i].done_today) habitsDoneCount++;
  }

  // 如果 Summary 还没今天的数据，补上 habits_done
  if (summary[todayISO]) {
    summary[todayISO].habits_done = summary[todayISO].habits_done || habitsDoneCount;
  }

  // ---- Grid (今天的 time slots) ----
  var grid = readTodayGrid_();

  // ---- Week start ----
  var weekStart = _weekStartISO(tz);

  return {
    day: todayISO,
    week_start: weekStart,
    summary: summary,
    quest: {
      habit: quest.habit,
      task: quest.task,
      habit_done: quest.habitDone,
      task_done: quest.taskDone,
      combo_days: quest.comboDays,
      combo_mult: quest.comboMultiplier
    },
    progress: {
      xp: xp.xp,
      level: xp.level,
      into_level: xp.intoLevel,
      level_span: xp.levelSpan,
      badges: xp.badges,
      boss: boss ? {
        id: '', name: boss.name, emoji: boss.emoji,
        target: boss.target, progress: boss.progress, defeated: boss.defeated
      } : null
    },
    habits: habits,
    grid: grid
  };
}

// ---- 习惯列表读取 ----

function readHabits_(todayDate) {
  var sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return [];

  var H = CONFIG.HABITS;
  var lastRow = getLastHabitRow_();
  if (lastRow < H.DATA_START_ROW) return [];

  var dataStart = H.DATA_START_ROW;
  var numHabits = lastRow - dataStart + 1;

  // 读条件 (A 列)、反应/名称 (B 列) 和分数 (D 列)
  var cueCol = columnLetterToIndex(H.COLUMNS.HABIT_CUE) + 1;      // A
  var nameCol = columnLetterToIndex(H.COLUMNS.HABIT_NAME) + 1;    // B
  var scoreCol = columnLetterToIndex(H.COLUMNS.WEIGHT) + 1;       // D
  var cues = sheet.getRange(dataStart, cueCol, numHabits, 1).getValues().flat();
  var names = sheet.getRange(dataStart, nameCol, numHabits, 1).getValues().flat();
  var scores = sheet.getRange(dataStart, scoreCol, numHabits, 1).getValues().flat();

  // 找今天在 14 天窗口中的位置
  var dayIndex = findHabitsDayColumnIndex_(sheet, todayDate);
  var dayStartCol = columnLetterToIndex(H.COLUMNS.DAY_START) + 1;  // E

  var habits = [];
  for (var i = 0; i < numHabits; i++) {
    var name = String(names[i]).trim();
    if (!name) continue;

    var doneToday = false;
    if (dayIndex >= 0) {
      // 今天的完成列 (E=5 → dayIndex=0)
      var doneCol = dayStartCol + dayIndex; // 1-based column index
      var cellVal = sheet.getRange(dataStart + i, doneCol).getValue();
      doneToday = !!(cellVal && parseFloat(cellVal) > 0);
    }

    // 计算连续天数（从今天往回数）
    var streakDays = _calcStreak(sheet, dataStart + i, dayIndex);

    habits.push({
      name: name,
      cue: String(cues[i] == null ? '' : cues[i]).trim(),
      weight: parseFloat(scores[i]) || 0,
      streak_days: streakDays,
      done_today: doneToday
    });
  }
  return habits;
}

/**
 * 返回今天在 14 天窗口中对应的列索引（0-based），今天不在窗口中返回 -1。
 */
function findHabitsDayColumnIndex_(sheet, todayDate) {
  var H = CONFIG.HABITS;
  var yearMonthCell = sheet.getRange(H.YEAR_MONTH_CELL).getValue();
  var firstDayVal = sheet.getRange(H.HEADER_RANGE.split(':')[0]).getValue();
  var firstDay = parseInt(firstDayVal, 10);
  if (!firstDay) return -1;

  var year, month;
  if (yearMonthCell instanceof Date) {
    year = yearMonthCell.getFullYear();
    month = yearMonthCell.getMonth();
  } else {
    var parts = String(yearMonthCell || '').trim().split(/\s+/);
    if (parts.length < 2) return -1;
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10) - 1;
  }
  if (isNaN(year) || isNaN(month)) return -1;

  var windowStart = new Date(year, month, firstDay);
  windowStart.setHours(0, 0, 0, 0);
  var t = new Date(todayDate);
  t.setHours(0, 0, 0, 0);

  var diff = (t - windowStart) / (1000 * 60 * 60 * 24);
  if (diff >= 0 && diff < H.DAYS_COUNT) return diff;
  return -1;
}

function _calcStreak(sheet, row, dayIndex) {
  if (dayIndex < 0) return 0;
  var start = columnLetterToIndex(CONFIG.HABITS.COLUMNS.DAY_START) + 1; // E
  var count = 0;
  for (var col = start + dayIndex; col >= start; col--) {
    var val = sheet.getRange(row, col).getValue();
    if (val && parseFloat(val) > 0) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

// ---- 今日 Grid 读取 ----

function readTodayGrid_() {
  var sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) return { slots: [], empty_slots: [] };

  var W = CONFIG.WEEKLY;
  var dayIndex = getCurrentDayIndex();
  if (dayIndex < 0 || dayIndex >= W.DAYS_IN_WEEK) {
    return { slots: [], empty_slots: [] };
  }

  var taskCol = W.TASK_COLUMNS[dayIndex];    // 1-based
  var scoreCol = W.SCORE_COLUMNS[dayIndex];  // 1-based

  var numRows = W.LAST_TIME_ROW - W.DATA_START_ROW + 1;

  // 读时间、任务、分数
  var timeVals = sheet.getRange(W.DATA_START_ROW, W.TIME_COLUMN, numRows, 1).getValues().flat();
  var taskVals = sheet.getRange(W.DATA_START_ROW, taskCol, numRows, 1).getValues().flat();
  var scoreVals = sheet.getRange(W.DATA_START_ROW, scoreCol, numRows, 1).getValues().flat();

  var slots = [];
  var emptySlots = [];

  for (var i = 0; i < numRows; i++) {
    var timeVal = timeVals[i];
    var taskVal = String(taskVals[i] || '').trim();
    var scoreVal = scoreVals[i];

    var timeStr = _decimalToTimeStr(timeVal);
    if (!timeStr) continue;

    if (taskVal) {
      var score = parseFloat(scoreVal);
      slots.push({
        time: timeStr,
        task: taskVal,
        score: (isFinite(score) && score > 0) ? score : 0
      });
    } else {
      // 空档：没有安排任务的时段
      // 只取 08:00-23:30 之间的空档
      var hour = _decimalHour(timeVal);
      if (hour >= 8 && hour < 24) {
        emptySlots.push(timeStr);
      }
    }
  }

  return { slots: slots, empty_slots: emptySlots };
}

function _decimalToTimeStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, getSpreadsheet_().getSpreadsheetTimeZone(), 'HH:mm');
  }
  var num = parseFloat(val);
  if (isNaN(num)) return null;
  var hour = Math.floor(num);
  var min = Math.round((num - hour) * 60);
  if (min === 60) { hour++; min = 0; }
  return String(hour).padStart(2, '0') + ':' + String(min).padStart(2, '0');
}

function _decimalHour(val) {
  if (val instanceof Date) return val.getHours() + val.getMinutes() / 60;
  var num = parseFloat(val);
  return isNaN(num) ? -1 : num;
}

function _weekStartISO(tz) {
  var monday = getMonday(new Date());
  return Utilities.formatDate(monday, tz, 'yyyy-MM-dd');
}
