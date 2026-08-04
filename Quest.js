/**
 * Daily Quest feature (Google Sheets edition).
 *
 * Each day one Habit and one Task are featured as the "Quest of the Day".
 * Completing/scoring a featured item grants a bonus multiplier
 * (CONFIG.QUEST.BONUS_MULTIPLIER) and both picks are highlighted in gold
 * on their sheets and shown in the sidebar.
 *
 * Selection is deterministic per day (seeded by the YYYYMMDD date) AND
 * persisted to DocumentProperties, so the same two picks are stable all
 * day across reloads/recalcs. A fresh pick happens only when the stored
 * date differs from today (driven by initialize()/dailyMaintenance()).
 *
 * Scoring hooks live in Habits.recordHabitDone and
 * Weekly.processWeeklyScoreChange, which call the questXxxMultiplier_ and
 * markQuestXxxDone_ helpers below.
 */

/**
 * Deterministic 0-based index for a (date, salt) pair over `n` items.
 * Pure function (FNV-1a hash) — same inputs always yield the same index,
 * which is what makes "same date -> same pick" verifiable.
 * @param {string} dateStr e.g. '20260615'
 * @param {string} salt independent stream per pick ('habit' vs 'task')
 * @param {number} n number of candidates
 * @returns {number} index in [0, n) or -1 when n <= 0
 */
function questIndexForDate_(dateStr, salt, n) {
  if (!n || n <= 0) return -1;
  let h = 2166136261; // FNV-1a 32-bit offset basis
  const s = String(dateStr) + '|' + String(salt);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h * 16777619) >>> 0; // * FNV prime, keep unsigned 32-bit
  }
  return h % n;
}

/**
 * Non-empty habit names with their 1-based rows, in sheet order.
 * @returns {Array<{name:string,row:number}>}
 */
function questHabitCandidates_() {
  const sheet = getSheetByName_(CONFIG.HABITS_SHEET);
  if (!sheet) return [];
  const start = CONFIG.HABITS.DATA_START_ROW;
  const last = getLastHabitRow_();
  if (last < start) return [];
  const vals = sheet.getRange(start, 1, last - start + 1, 1).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const name = vals[i][0];
    if (name !== '' && name !== null) out.push({ name: String(name), row: start + i });
  }
  return out;
}

/**
 * Non-empty task names eligible to be featured: positive weight only
 * (negative-weight tasks are things to avoid, not quests) and excluding
 * the 'others' fallback. Returns names with their 1-based rows.
 * @returns {Array<{name:string,row:number}>}
 */
function questTaskCandidates_() {
  const sheet = getSheetByName_(CONFIG.TASKS_SHEET);
  if (!sheet) return [];
  const start = CONFIG.TASKS.DATA_START_ROW;
  const last = getLastTaskRow_();
  if (last < start) return [];
  const rows = sheet.getRange(start, 1, last - start + 1, 2).getValues(); // A=name, B=weight
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const name = rows[i][0];
    if (name === '' || name === null) continue;
    if (String(name) === CONFIG.TASKS.FALLBACK_NAME) continue;
    const parsed = parseFloat(rows[i][1]);
    const weight = isFinite(parsed) ? parsed : 1; // blank weight defaults to 1
    if (weight > 0) out.push({ name: String(name), row: start + i });
  }
  return out;
}

/**
 * Parse the persisted quest, or null.
 * @returns {{date:string,habitName:string,taskName:string,habitDone:boolean,taskDone:boolean}|null}
 */
function getDailyQuest_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.QUEST.PROP_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

/**
 * Persist the quest object.
 * @param {Object} quest
 */
function saveDailyQuest_(quest) {
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG.QUEST.PROP_KEY,
    JSON.stringify(quest)
  );
}

/**
 * Ensure today's quest exists. Re-picks (and refreshes highlights) only
 * when the stored date != today; otherwise returns the persisted quest
 * unchanged so picks stay stable all day. Safe to call repeatedly.
 * @returns {{date:string,habitName:string,taskName:string,habitDone:boolean,taskDone:boolean}}
 */
function ensureDailyQuest_() {
  const today = formatDateYYYYMMDD(new Date());
  const existing = getDailyQuest_();
  if (existing && existing.date === today) return existing;

  const habits = questHabitCandidates_();
  const tasks = questTaskCandidates_();
  const hIdx = questIndexForDate_(today, 'habit', habits.length);
  const tIdx = questIndexForDate_(today, 'task', tasks.length);

  const quest = {
    date: today,
    habitName: hIdx >= 0 ? habits[hIdx].name : '',
    taskName: tIdx >= 0 ? tasks[tIdx].name : '',
    habitDone: false,
    taskDone: false,
  };

  clearQuestHighlights_(existing); // remove yesterday's gold marks
  saveDailyQuest_(quest);
  applyQuestHighlights_(quest);
  return quest;
}

/**
 * 1-based row of a habit by name, or -1.
 * @param {string} name
 * @returns {number}
 */
function questHabitRowByName_(name) {
  if (!name) return -1;
  const found = questHabitCandidates_().filter((c) => c.name === String(name));
  return found.length ? found[0].row : -1;
}

/**
 * 1-based row of a task by name, or -1.
 * @param {string} name
 * @returns {number}
 */
function questTaskRowByName_(name) {
  if (!name) return -1;
  const found = questTaskCandidates_().filter((c) => c.name === String(name));
  return found.length ? found[0].row : -1;
}

/**
 * Paint the featured habit/task name cells gold. Best-effort.
 * @param {Object|null} quest
 */
function applyQuestHighlights_(quest) {
  if (!quest) return;
  safeInit_('Quest highlight (habit) failed', function () {
    const row = questHabitRowByName_(quest.habitName);
    if (row > 0) {
      getSheetByName_(CONFIG.HABITS_SHEET)
        .getRange(row, 1)
        .setBackground(CONFIG.COLORS.QUEST_HIGHLIGHT);
    }
  });
  safeInit_('Quest highlight (task) failed', function () {
    const row = questTaskRowByName_(quest.taskName);
    if (row > 0) {
      getSheetByName_(CONFIG.TASKS_SHEET)
        .getRange(row, 1)
        .setBackground(CONFIG.COLORS.QUEST_HIGHLIGHT);
    }
  });
}

/**
 * Clear the gold marks from a previous quest's featured cells. Best-effort.
 * @param {Object|null} quest
 */
function clearQuestHighlights_(quest) {
  if (!quest) return;
  safeInit_('Quest unhighlight (habit) failed', function () {
    const row = questHabitRowByName_(quest.habitName);
    if (row > 0) {
      getSheetByName_(CONFIG.HABITS_SHEET).getRange(row, 1).setBackground(CONFIG.COLORS.CLEAR);
    }
  });
  safeInit_('Quest unhighlight (task) failed', function () {
    const row = questTaskRowByName_(quest.taskName);
    if (row > 0) {
      getSheetByName_(CONFIG.TASKS_SHEET).getRange(row, 1).setBackground(CONFIG.COLORS.CLEAR);
    }
  });
}

/**
 * Bonus multiplier to apply when the given habit is today's featured habit.
 * @param {string} habitName
 * @returns {number} CONFIG.QUEST.BONUS_MULTIPLIER or 1
 */
function questHabitMultiplier_(habitName) {
  const q = getDailyQuest_();
  return q && habitName && q.habitName === String(habitName)
    ? CONFIG.QUEST.BONUS_MULTIPLIER
    : 1;
}

/**
 * Bonus multiplier to apply when the given task is today's featured task.
 * @param {string} taskName
 * @returns {number} CONFIG.QUEST.BONUS_MULTIPLIER or 1
 */
function questTaskMultiplier_(taskName) {
  const q = getDailyQuest_();
  return q && taskName && q.taskName === String(taskName)
    ? CONFIG.QUEST.BONUS_MULTIPLIER
    : 1;
}

/**
 * Mark the quest's habitDone/taskDone flag. Locked read-modify-write so a
 * habit completion and a task score landing together can't lose a flag.
 * @param {'habit'|'task'} which
 * @param {string} name only marks when it matches the featured item
 */
function markQuestDone_(which, name) {
  if (!name) return;
  let lock = null;
  try {
    lock = LockService.getDocumentLock();
    lock.tryLock(2000);
  } catch (e) {
    lock = null;
  }
  let flipped = false;
  try {
    const q = getDailyQuest_();
    if (!q) return;
    if (which === 'habit' && q.habitName === String(name) && !q.habitDone) {
      q.habitDone = true;
      saveDailyQuest_(q);
      flipped = true;
    } else if (which === 'task' && q.taskName === String(name) && !q.taskDone) {
      q.taskDone = true;
      saveDailyQuest_(q);
      flipped = true;
    }
  } catch (e) {
    Logger.log('markQuestDone_ failed: ' + (e && e.message ? e.message : e));
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        // ignore
      }
    }
  }
  // Count toward the Quest Master badge + weekly boss (after our lock).
  if (flipped) {
    recordQuestCompletion_();
    bumpBossQuestCount_();
  }
}

/**
 * Sidebar-callable view of today's quest. Self-heals (picks if missing).
 * @returns {{habit:string,task:string,habitDone:boolean,taskDone:boolean,
 *   bonus:number,comboDays:number,comboMultiplier:number,
 *   comboDoneToday:boolean,comboAtRisk:boolean}}
 */
function getDailyQuestFromUI() {
  const q = ensureDailyQuest_();
  const combo = getComboState_();
  return {
    habit: q.habitName || '',
    task: q.taskName || '',
    habitDone: !!q.habitDone,
    taskDone: !!q.taskDone,
    bonus: CONFIG.QUEST.BONUS_MULTIPLIER,
    comboDays: combo.days,
    comboMultiplier: combo.multiplier,
    comboDoneToday: combo.doneToday,
    comboAtRisk: combo.atRisk,
  };
}

// ----------------------------------------------------------------------
// Quest streak combos
// ----------------------------------------------------------------------
// "Completing the quest" for a day = the featured habit OR task was done
// that day (either counts). Consecutive completed days build a combo
// multiplier applied on top of the quest bonus; a missed day resets it.

/**
 * Multiplier for a given consecutive-day count. Pure + testable.
 * 0 days => BASE (1.0, no combo); n days => BASE + min(n, CAP)*STEP.
 * @param {number} days
 * @param {number} base
 * @param {number} step
 * @param {number} capDays
 * @returns {number}
 */
function comboMultiplier_(days, base, step, capDays) {
  if (!days || days <= 0) return base;
  return base + Math.min(days, capDays) * step;
}

/** @returns {{combo:number, lastDate:string}} */
function getQuestCombo_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.COMBO.PROP_KEY);
    if (raw) {
      const o = JSON.parse(raw);
      return { combo: Number(o.combo) || 0, lastDate: String(o.lastDate || '') };
    }
  } catch (e) {
    // fall through
  }
  return { combo: 0, lastDate: '' };
}

/** @param {{combo:number, lastDate:string}} state */
function saveQuestCombo_(state) {
  PropertiesService.getDocumentProperties().setProperty(CONFIG.COMBO.PROP_KEY, JSON.stringify(state));
}

/** @returns {string} yesterday's YYYYMMDD */
function questYesterday_() {
  const y = new Date();
  y.setDate(y.getDate() - 1);
  return formatDateYYYYMMDD(y);
}

/**
 * Read-only: the combo multiplier that applies to a quest completion made
 * NOW (uses combo+1 when the streak is alive from yesterday, or the locked
 * value if already counted today). No side effects — safe inside a lock.
 * @returns {number}
 */
function comboMultiplierForToday_() {
  const C = CONFIG.COMBO;
  const s = getQuestCombo_();
  const today = formatDateYYYYMMDD(new Date());
  let n;
  if (s.lastDate === today) n = s.combo; // already counted today
  else if (s.lastDate === questYesterday_()) n = s.combo + 1; // continues
  else n = 1; // gap or first
  return comboMultiplier_(n, C.BASE, C.STEP, C.CAP_DAYS);
}

/**
 * Commit today's combo (idempotent per day). Call AFTER releasing the
 * caller's document lock (this takes its own lock).
 * @returns {number} the combo day-count after advancing
 */
function advanceComboForToday_() {
  let result = 0;
  let lock = null;
  try {
    lock = LockService.getDocumentLock();
    lock.tryLock(2000);
  } catch (e) {
    lock = null;
  }
  try {
    const s = getQuestCombo_();
    const today = formatDateYYYYMMDD(new Date());
    let combo;
    if (s.lastDate === today) combo = s.combo; // already counted today
    else if (s.lastDate === questYesterday_()) combo = s.combo + 1;
    else combo = 1;
    saveQuestCombo_({ combo: combo, lastDate: today });
    result = combo;
  } catch (e) {
    Logger.log('advanceComboForToday_ failed: ' + (e && e.message ? e.message : e));
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        // ignore
      }
    }
  }
  return result;
}

/**
 * Display state for the sidebar / email: current alive streak, the
 * multiplier that applies to a completion today, and whether today is
 * done or the streak is at risk.
 * @returns {{days:number, multiplier:number, doneToday:boolean, atRisk:boolean}}
 */
function getComboState_() {
  const C = CONFIG.COMBO;
  const s = getQuestCombo_();
  const today = formatDateYYYYMMDD(new Date());
  if (s.lastDate === today) {
    return {
      days: s.combo,
      multiplier: comboMultiplier_(s.combo, C.BASE, C.STEP, C.CAP_DAYS),
      doneToday: true,
      atRisk: false,
    };
  }
  if (s.lastDate === questYesterday_()) {
    // Alive from yesterday but not yet done today — finishing extends it.
    return {
      days: s.combo,
      multiplier: comboMultiplier_(s.combo + 1, C.BASE, C.STEP, C.CAP_DAYS),
      doneToday: false,
      atRisk: true,
    };
  }
  // Broken / none — finishing today starts a fresh combo.
  return {
    days: 0,
    multiplier: comboMultiplier_(1, C.BASE, C.STEP, C.CAP_DAYS),
    doneToday: false,
    atRisk: false,
  };
}

/**
 * Click handler for the sidebar Quest "Habit" row: mark today's featured
 * habit done (same path as ticking its checkbox, so the streak + quest
 * bonus apply).
 * @returns {string} status message
 */
function completeQuestHabitFromUI() {
  const q = ensureDailyQuest_();
  if (!q || !q.habitName) return 'No quest habit today.';
  const row = questHabitRowByName_(q.habitName);
  if (row < 0) return 'Quest habit "' + q.habitName + '" not found on the Habits sheet.';
  recordHabitDone(row);
  return 'Marked quest habit "' + q.habitName + '" done.';
}

/**
 * Click handler for the sidebar Quest "Task" row: drop today's featured
 * task into the current time slot of today on the Weekly sheet (overriding
 * whatever was there) and auto-score it 1.0 — which colors the cell green,
 * updates the daily total / Summary / Tasks stats, applies the quest
 * bonus, and marks the quest task done.
 * @returns {string} status message
 */
function scheduleQuestTaskFromUI() {
  const q = ensureDailyQuest_();
  if (!q || !q.taskName) return 'No quest task today.';
  if (q.taskDone) return 'Quest task "' + q.taskName + '" already done today.';

  const sheet = getSheetByName_(CONFIG.WEEKLY_SHEET);
  if (!sheet) return 'Weekly sheet not found.';

  const row = getCurrentTimeRow_(sheet);
  if (row < 0) return 'No current time slot found on the Weekly sheet.';

  const dayIndex = getCurrentDayIndex();
  const taskCol = getTaskColForDay(dayIndex);
  const scoreCol = getScoreColForDay(dayIndex);

  // Don't clobber / double-count a slot that already has a score.
  const existingScore = sheet.getRange(row, scoreCol).getValue();
  if (existingScore !== '' && existingScore !== null) {
    return 'Current time slot is already scored — skipped.';
  }

  sheet.getRange(row, taskCol).setValue(q.taskName);
  // Auto-score a full 1.0; processWeeklyScoreChange handles the green
  // color, totals, Summary, quest bonus, and marking the quest done.
  processWeeklyScoreChange(row, scoreCol, 1);

  return 'Completed quest task "' + q.taskName + '" in the current slot.';
}
