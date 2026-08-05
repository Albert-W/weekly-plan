/**
 * XP, Levels & Badges (Google Sheets edition).
 *
 * A persistent progression layer on top of scoring. Every positive point
 * scored (weekly grid score incl. quest bonus, habit completion) also
 * grants XP, stored in DocumentProperties so it survives New Week resets.
 * Crossing a level threshold toasts a level-up; achievements unlock
 * one-time badges.
 *
 * Scoring hooks (Habits.recordHabitDone, Weekly.processWeeklyScoreChange)
 * call awardXp_ after their locks release; Quest.markQuestDone_ calls
 * recordQuestCompletion_.
 */

/** Badge id -> display metadata. */
const BADGE_META_ = {
  centurion: { emoji: '💯', label: 'Centurion', desc: 'Earned 100 lifetime XP' },
  rising_star: { emoji: '🌟', label: 'Rising Star', desc: 'Reached Level 5' },
  week_warrior: { emoji: '🔥', label: 'Week Warrior', desc: '7-day habit streak' },
  quest_master: { emoji: '🎯', label: 'Quest Master', desc: 'Completed 10 quest items' },
  early_bird: { emoji: '🌅', label: 'Early Bird', desc: 'Logged a win before 9am' },
  boss_slayer: { emoji: '🐉', label: 'Boss Slayer', desc: 'Defeated a Weekly Boss' },
};

/**
 * Level + the XP bounds of that level for a given lifetime XP total.
 * Pure + testable. Threshold to advance level L is BASE + (L-1)*STEP.
 * @param {number} totalXp
 * @param {number} base
 * @param {number} step
 * @returns {{level:number, levelFloor:number, nextThreshold:number}}
 */
function levelForXp_(totalXp, base, step) {
  const xp = totalXp > 0 ? totalXp : 0;
  let level = 1;
  let floor = 0;
  let span = base;
  while (xp >= floor + span) {
    floor += span;
    level++;
    span += step;
  }
  return { level: level, levelFloor: floor, nextThreshold: floor + span };
}

/**
 * Parsed XP state, defaulting to a fresh level-1 record.
 * @returns {{xp:number, level:number}}
 */
function getXpState_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.XP.PROP_KEY);
    if (raw) {
      const o = JSON.parse(raw);
      return { xp: Number(o.xp) || 0, level: Number(o.level) || 1 };
    }
  } catch (e) {
    // fall through to default
  }
  return { xp: 0, level: 1 };
}

/**
 * Persist XP state.
 * @param {{xp:number, level:number}} state
 */
function saveXpState_(state) {
  PropertiesService.getDocumentProperties().setProperty(CONFIG.XP.PROP_KEY, JSON.stringify(state));
}

/** @returns {string[]} earned badge ids */
function getEarnedBadgeIds_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.XP.BADGES_PROP);
    return raw ? JSON.parse(raw) || [] : [];
  } catch (e) {
    return [];
  }
}

/**
 * Award a badge once (idempotent). Toasts on first unlock.
 * @param {string} id key in BADGE_META_
 */
function awardBadge_(id) {
  if (!BADGE_META_[id]) return;
  const awarded = tryWithLock_(function () {
    try {
      const ids = getEarnedBadgeIds_();
      if (ids.indexOf(id) === -1) {
        ids.push(id);
        PropertiesService.getDocumentProperties().setProperty(
          CONFIG.XP.BADGES_PROP,
          JSON.stringify(ids)
        );
        return true;
      }
      return false;
    } catch (e) {
      Logger.log('awardBadge_ failed: ' + (e && e.message ? e.message : e));
      return false;
    }
  });
  if (awarded) {
    const m = BADGE_META_[id];
    toast_('🏅 Badge unlocked: ' + m.emoji + ' ' + m.label, 'Weekly Plan', 'success');
  }
}

/**
 * Add XP (positive only), detect level-ups, and check XP/level badges.
 * Lifetime — never reset by New Week.
 * @param {number} points
 */
function awardXp_(points) {
  if (!points || points <= 0) return;

  let oldLevel = 1;
  let newXp = 0;
  let newLevel = 1;

  const result = tryWithLock_(function () {
    try {
      const s = getXpState_();
      oldLevel = s.level || 1;
      newXp = (s.xp || 0) + points;
      newLevel = levelForXp_(newXp, CONFIG.XP.BASE, CONFIG.XP.STEP).level;
      saveXpState_({ xp: newXp, level: newLevel });
      return true;
    } catch (e) {
      Logger.log('awardXp_ failed: ' + (e && e.message ? e.message : e));
      return false;
    }
  });

  if (!result) return;

  if (newLevel > oldLevel) {
    toast_('⬆️ Level up! You reached Level ' + newLevel + '.', 'Weekly Plan', 'success');
  }
  if (newXp >= CONFIG.XP.CENTURION_XP) awardBadge_('centurion');
  if (newLevel >= CONFIG.XP.RISING_STAR_LEVEL) awardBadge_('rising_star');
}

/**
 * Record one completed quest item (habit or task) toward the Quest Master
 * badge. Called from Quest.markQuestDone_ when a flag flips.
 */
function recordQuestCompletion_() {
  let count = 0;
  tryWithLock_(function () {
    try {
      const props = PropertiesService.getDocumentProperties();
      count = (parseInt(props.getProperty(CONFIG.XP.QUEST_COUNT_PROP), 10) || 0) + 1;
      props.setProperty(CONFIG.XP.QUEST_COUNT_PROP, String(count));
    } catch (e) {
      Logger.log('recordQuestCompletion_ failed: ' + (e && e.message ? e.message : e));
    }
  });
  if (count >= CONFIG.XP.QUEST_MASTER_COUNT) awardBadge_('quest_master');
}

/**
 * Award the Week Warrior badge when a habit streak reaches the threshold.
 * @param {number} streakDays the streak length achieved (incl. today)
 */
function maybeAwardStreakBadge_(streakDays) {
  if (streakDays >= CONFIG.XP.WEEK_WARRIOR_STREAK) awardBadge_('week_warrior');
}

/**
 * Award the Early Bird badge for a win logged before EARLY_BIRD_HOUR.
 * @param {Date} when
 */
function maybeAwardEarlyBird_(when) {
  if (when.getHours() < CONFIG.XP.EARLY_BIRD_HOUR) awardBadge_('early_bird');
}

/**
 * Sidebar-callable XP/level/badge snapshot.
 * @returns {{level:number, xp:number, intoLevel:number, levelSpan:number,
 *   nextThreshold:number, progressPct:number,
 *   badges:Array<{emoji:string,label:string,desc:string}>}}
 */
function getXpStateFromUI() {
  const s = getXpState_();
  const info = levelForXp_(s.xp, CONFIG.XP.BASE, CONFIG.XP.STEP);
  const span = info.nextThreshold - info.levelFloor;
  const into = s.xp - info.levelFloor;
  const round2 = (n) => Math.round(n * 100) / 100;
  const badges = getEarnedBadgeIds_()
    .filter((id) => BADGE_META_[id])
    .map((id) => BADGE_META_[id]);
  return {
    level: info.level,
    xp: round2(s.xp),
    intoLevel: round2(into),
    levelSpan: round2(span),
    nextThreshold: round2(info.nextThreshold),
    progressPct: span > 0 ? Math.round((into / span) * 100) : 0,
    badges: badges,
  };
}
