/**
 * Weekly Boss challenge (Google Sheets edition).
 *
 * One rotating objective per week, deterministically chosen from
 * CONFIG.BOSS.DEFS (so it varies week to week but is stable within a
 * week). Defeating it (progress >= target) grants REWARD_XP and the
 * 🐉 Boss Slayer badge, exactly once per week. A fresh boss is selected
 * when the ISO week (Monday) changes.
 *
 * Progress sources:
 *   - 'points'  : this week's Summary total (sum of daily totals).
 *   - 'quests'  : quest items completed this week (counter on boss state).
 *
 * Hooks: scoring paths (Habits.recordHabitDone, Weekly.processWeeklyScore-
 * Change) call checkBossDefeat_ after awarding XP; Quest.markQuestDone_
 * calls bumpBossQuestCount_.
 */

/**
 * Boss definition by id, or null.
 * @param {string} id
 * @returns {Object|null}
 */
function bossDefById_(id) {
  const defs = CONFIG.BOSS.DEFS || [];
  for (let i = 0; i < defs.length; i++) {
    if (defs[i].id === id) return defs[i];
  }
  return null;
}

/** @returns {{weekStart:string, bossId:string, defeated:boolean, questCount:number}} */
function getBossState_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.BOSS.PROP_KEY);
    if (raw) {
      const o = JSON.parse(raw);
      return {
        weekStart: String(o.weekStart || ''),
        bossId: String(o.bossId || ''),
        defeated: !!o.defeated,
        questCount: Number(o.questCount) || 0,
      };
    }
  } catch (e) {
    // fall through
  }
  return { weekStart: '', bossId: '', defeated: false, questCount: 0 };
}

/** @param {Object} state */
function saveBossState_(state) {
  PropertiesService.getDocumentProperties().setProperty(CONFIG.BOSS.PROP_KEY, JSON.stringify(state));
}

/**
 * Ensure the current week's boss is selected; re-pick (resetting progress)
 * when the week changes. Deterministic per week.
 * @returns {Object} boss state
 */
function ensureWeeklyBoss_() {
  const weekStart = formatDateYYYYMMDD(getMonday(new Date()));
  const s = getBossState_();
  if (s.weekStart === weekStart && s.bossId) return s;

  const defs = CONFIG.BOSS.DEFS || [];
  const idx = defs.length ? questIndexForDate_(weekStart, 'boss', defs.length) : -1;
  const boss = {
    weekStart: weekStart,
    bossId: idx >= 0 ? defs[idx].id : '',
    defeated: false,
    questCount: 0,
  };
  saveBossState_(boss);
  return boss;
}

/**
 * Sum this week's (Mon–Sun) Summary totals.
 * @returns {number}
 */
function weeklyPoints_() {
  const sheet = getSheetByName_(CONFIG.SUMMARY_SHEET);
  if (!sheet) return 0;
  const monday = getMonday(new Date());
  const start = formatDateYYYYMMDD(monday);
  const endDate = new Date(monday);
  endDate.setDate(monday.getDate() + 6);
  const end = formatDateYYYYMMDD(endDate);

  const lastRow = Math.max(getLastRowInColumn_(sheet, 1), 1);
  const rows = sheet.getRange('A1:F' + lastRow).getValues();
  let sum = 0;
  for (let i = 0; i < rows.length; i++) {
    const d = String(rows[i][0]);
    if (d >= start && d <= end) {
      const t = parseFloat(rows[i][5]);
      if (isFinite(t)) sum += t;
    }
  }
  return sum;
}

/**
 * Current progress toward the active boss.
 * @param {Object} state boss state
 * @param {Object} def boss definition
 * @returns {number}
 */
function bossProgress_(state, def) {
  if (!def) return 0;
  if (def.type === 'quests') return state.questCount || 0;
  return weeklyPoints_(); // 'points'
}

/**
 * Check whether the active boss is defeated; if newly defeated, award XP +
 * badge exactly once. Safe to call after any scoring/quest event.
 */
function checkBossDefeat_() {
  const state = ensureWeeklyBoss_();
  const def = bossDefById_(state.bossId);
  if (!def || state.defeated) return;
  if (bossProgress_(state, def) < def.target) return;

  const newlyDefeated = tryWithLock_(function () {
    try {
      const s = getBossState_();
      if (s.weekStart === state.weekStart && s.bossId === state.bossId && !s.defeated) {
        s.defeated = true;
        saveBossState_(s);
        return true;
      }
      return false;
    } catch (e) {
      Logger.log('checkBossDefeat_ failed: ' + (e && e.message ? e.message : e));
      return false;
    }
  });

  if (newlyDefeated) {
    awardXp_(CONFIG.BOSS.REWARD_XP);
    awardBadge_('boss_slayer');
    toast_(
      def.emoji + ' Boss defeated: ' + def.name + '! +' + CONFIG.BOSS.REWARD_XP + ' XP',
      'Weekly Plan',
      'success'
    );
  }
}

/**
 * Count one quest item toward a quest-type boss, then re-check defeat.
 * Called from Quest.markQuestDone_ when a flag flips.
 */
function bumpBossQuestCount_() {
  tryWithLock_(function () {
    try {
      const s = ensureWeeklyBoss_();
      s.questCount = (s.questCount || 0) + 1;
      saveBossState_(s);
    } catch (e) {
      Logger.log('bumpBossQuestCount_ failed: ' + (e && e.message ? e.message : e));
    }
  });
  checkBossDefeat_();
}

/**
 * Sidebar/email-callable snapshot of the current boss.
 * @returns {{name:string, emoji:string, type:string, target:number,
 *   progress:number, remaining:number, pct:number, defeated:boolean}|null}
 */
function getBossStateFromUI() {
  const s = ensureWeeklyBoss_();
  const def = bossDefById_(s.bossId);
  if (!def) return null;
  const progress = bossProgress_(s, def);
  const round1 = (n) => Math.round(n * 10) / 10;
  const pct = def.target > 0 ? Math.min(100, Math.max(0, Math.round((progress / def.target) * 100))) : 0;
  return {
    name: def.name,
    emoji: def.emoji,
    type: def.type,
    target: def.target,
    progress: round1(progress),
    remaining: round1(Math.max(0, def.target - progress)),
    pct: pct,
    defeated: !!s.defeated,
  };
}
