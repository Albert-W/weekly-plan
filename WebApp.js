/**
 * Web App 入口 — 部署为 GAS Web App（执行身份：我，访问权限：任何人）。
 *
 * GET  ?k=<auth> → 返回每日快照 JSON（供 Mac 定时拉取）
 * POST ?k=<auth> + JSON body → 执行操作（习惯打卡等）
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

// ---- GET: 每日快照 ----

function doGet(e) {
  if (!auth_(e)) return unauth_();
  try {
    var snapshot = buildSnapshot_();
    snapshot.ok = true;
    return json_(snapshot);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ---- POST: 操作入口 ----

function doPost(e) {
  if (!auth_(e)) return unauth_();
  try {
    var params = JSON.parse(e.postData.contents);
    // { action: "habit", name: "健身" }
    if (params.action === 'habit' && params.name) {
      return handleHabitCheckin_(params.name);
    }
    return json_({ ok: false, error: 'unknown action: ' + (params.action || '') });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ---- 习惯打卡 ----

function handleHabitCheckin_(name) {
  var sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return json_({ ok: false, error: 'Habits sheet not found' });

  var lastRow = getLastHabitRow_();
  var names = sheet.getRange(CONFIG.HABITS.DATA_START_ROW, 1, lastRow - CONFIG.HABITS.DATA_START_ROW + 1, 1)
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

  // 读名称 (A 列) 和分数 (C 列)
  var nameRange = sheet.getRange(dataStart, 1, numHabits, 1);
  var scoreRange = sheet.getRange(dataStart, 3, numHabits, 1);
  var names = nameRange.getValues().flat();
  var scores = scoreRange.getValues().flat();

  // 找今天在 14 天窗口中的位置
  var dayIndex = findHabitsDayColumnIndex_(sheet, todayDate);

  var habits = [];
  for (var i = 0; i < numHabits; i++) {
    var name = String(names[i]).trim();
    if (!name) continue;

    var doneToday = false;
    if (dayIndex >= 0) {
      // 今天的完成列 (D=4 → dayIndex=0)
      var doneCol = 4 + dayIndex; // 1-based column index
      var cellVal = sheet.getRange(dataStart + i, doneCol).getValue();
      doneToday = !!(cellVal && parseFloat(cellVal) > 0);
    }

    // 计算连续天数（从今天往回数）
    var streakDays = _calcStreak(sheet, dataStart + i, dayIndex);

    habits.push({
      name: name,
      base_score: parseFloat(scores[i]) || 0,
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
  var count = 0;
  for (var col = 4 + dayIndex; col >= 4; col--) {
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
