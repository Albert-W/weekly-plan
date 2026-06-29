/**
 * Configuration constants for the Weekly Plan Google Sheets edition.
 *
 * Ported from src/taskpane/js/config.js. In Apps Script every .gs file
 * shares one global scope, so this top-level `const CONFIG` is visible
 * to every other module — no `window.*` export needed.
 */

const CONFIG = {
  // ==================== SHEET NAMES ====================
  HABITS_SHEET: 'Habits',
  WEEKLY_SHEET: 'Weekly',
  TASKS_SHEET: 'Tasks',
  SUMMARY_SHEET: 'Summary',
  ARCHIVE_SHEET: 'Archive', // GAS-only: in-spreadsheet history of finished weeks

  // ==================== HABITS CONFIG ====================
  HABITS: {
    DATA_START_ROW: 4,
    HEADER_ROW: 3,
    COLUMNS: {
      HABIT_NAME: 'A',
      DONE_CHECKBOX: 'B', // GAS-only: native checkbox column ("mark done")
      BASE_SCORE: 'C',
      DAY_START: 'D',
      DAY_END: 'Q',
      TOTAL_COUNT: 'R',
    },
    YEAR_MONTH_CELL: 'B3',
    HEADER_RANGE: 'D3:Q3',
    DAYS_COUNT: 14,
    STREAK_MULTIPLIER: 1.1,
  },

  // ==================== TASKS CONFIG ====================
  TASKS: {
    // Name used for the auto-created "catch-all" row when a scored
    // task isn't found in the Tasks list.
    FALLBACK_NAME: 'others',
    // First row of actual task data (rows 1-3 are headers).
    DATA_START_ROW: 4,
  },

  // ==================== WEEKLY/TIMETABLE CONFIG ====================
  WEEKLY: {
    DATA_START_ROW: 5,
    CONTROL_ROW: 2,
    HEADER_ROW: 4,
    TIME_COLUMN: 2, // Column B for timestamps
    DATE_CELL: 'B4', // Cell containing "yyyy mm" format
    FIRST_DAY_HEADER_CELL: 'D4', // First day-number header (Monday)
    HEADER_RANGE: 'D4:P4', // All 7 day-number headers (D, F, H, J, L, N, P)
    HEADER_ROW_RANGE: 'A4:P4', // Whole header row, for fill clears
    LAST_TIME_ROW: 36, // Last row with time data
    SCORE_ROW: 38, // Score totals row
    // Task columns (odd: 3,5,7,9,11,13,15) = C,E,G,I,K,M,O
    TASK_COLUMNS: [3, 5, 7, 9, 11, 13, 15],
    // Score columns (even: 4,6,8,10,12,14,16) = D,F,H,J,L,N,P
    SCORE_COLUMNS: [4, 6, 8, 10, 12, 14, 16],
    // Score options for the data-validation dropdown
    SCORE_OPTIONS: [0, 0.2, 0.4, 0.6, 0.8, 1],
    // Days in week
    DAYS_IN_WEEK: 7,
    // Scaffold: first hour and last hour for the time-block grid.
    // Rows DATA_START_ROW..LAST_TIME_ROW hold one 30-min slot each,
    // shown as decimal hours (8, 8.5, 9, …) to match the original.
    FIRST_HOUR: 8, // 08:00
    // Mid-day visual divider drawn under this decimal hour (e.g. 17.5).
    MID_DIVIDER_HOUR: 17.5,
    // In-sheet control buttons in CONTROL_ROW. Each is merged across
    // `col`..`col+1`. `action` maps to a handler in handleSelection (Triggers).
    CONTROL_BUTTONS: [
      { action: 'help', label: 'Help', col: 3 }, // C2:D2
      { action: 'add', label: 'Add Task', col: 5 }, // E2:F2
      { action: 'delete', label: 'Delete Task', col: 7 }, // G2:H2
      { action: 'random', label: 'Random Fill', col: 9 }, // I2:J2
      { action: 'thanks', label: 'Thanks', col: 11 }, // K2:L2
    ],
  },

  // ==================== SUMMARY CONFIG ====================
  SUMMARY: {
    DATE_COLUMN: 'A',
    POSITIVE_SCORE_COLUMN: 'D',
    NEGATIVE_SCORE_COLUMN: 'E',
    TOTAL_SCORE_COLUMN: 'F',
  },

  // ==================== DAILY QUEST CONFIG ====================
  QUEST: {
    // Featured habit/task earn this multiplier on their points when
    // completed/scored on their quest day (x1.5 => +50% bonus).
    BONUS_MULTIPLIER: 1.5,
    // DocumentProperties key holding the persisted daily quest JSON.
    PROP_KEY: 'dailyQuest',
  },

  // ==================== QUEST STREAK COMBO CONFIG ====================
  // Completing the daily quest (featured habit OR task done) on
  // consecutive days builds a combo multiplier applied ON TOP of the
  // quest bonus. Missing a day resets it. multiplier(n days) =
  // BASE + min(n, CAP_DAYS) * STEP  =>  1d x1.2, 2d x1.4 ... 5d+ x2.0.
  COMBO: {
    BASE: 1.0,
    STEP: 0.2,
    CAP_DAYS: 5,
    PROP_KEY: 'questCombo', // JSON { combo:number, lastDate:'YYYYMMDD' }
  },

  // ==================== MORNING EMAIL CONFIG ====================
  EMAIL: {
    // Master on/off switch for the daily summary email. Disabled by
    // default — the morning recap is delivered via Telegram instead (see
    // TELEGRAM below). Flip back to true to re-enable the email.
    ENABLED: false,
    // Recipient address. Empty string => the script owner's account
    // (Session.getEffectiveUser().getEmail()).
    RECIPIENT: '',
    // DocumentProperties key tracking the last date we sent (no dupes).
    LAST_SENT_PROP: 'lastEmailDate',
  },

  // ==================== MORNING TELEGRAM CONFIG ====================
  // Daily morning recap delivered via a Telegram bot (yesterday's score,
  // today's quest habit+task, combo, XP/level, weekly boss). Driven by
  // the time-based `morningTelegram` trigger (~SEND_HOUR) and deduped
  // per day via LAST_SENT_PROP. Requires the script.external_request
  // OAuth scope (see appsscript.json).
  TELEGRAM: {
    // Master on/off switch for the daily Telegram recap.
    ENABLED: true,
    // Bot credentials are NOT stored in source. They live in Script
    // Properties under these keys — set them once via the "Set up
    // Telegram…" menu item (setUpTelegram) or the Script Properties
    // editor. Token comes from @BotFather; chat id is your numeric id.
    BOT_TOKEN_PROP: 'telegramBotToken',
    CHAT_ID_PROP: 'telegramChatId',
    // Hour of day (0–23, spreadsheet timezone) for the morning send.
    // Apps Script time triggers fire within that hour window.
    SEND_HOUR: 8,
    // DocumentProperties key tracking the last date we sent (no dupes).
    LAST_SENT_PROP: 'lastTelegramDate',
  },

  // ==================== GEMINI (LLM) CONFIG ====================
  // Google Gemini is used to write the meal-time habit stories (see STORY
  // below). The API key lives in Script Properties (set via "Set up
  // Gemini…" / setUpGemini), never in source. Calls go out over HTTPS, so
  // the script.external_request OAuth scope is required.
  GEMINI: {
    API_KEY_PROP: 'geminiApiKey',
    // Model name on the Generative Language API. Change if a newer/older
    // model is preferred (e.g. 'gemini-1.5-flash').
    MODEL: 'gemini-3.1-flash-lite',
    // Higher temperature => more varied, surprising stories.
    TEMPERATURE: 1.0,
    MAX_OUTPUT_TOKENS: 500,
  },

  // ==================== MEAL-TIME HABIT STORY CONFIG ====================
  // Three times a day (meal times) a random habit is picked and Gemini
  // writes a short, motivating story whose heart is that habit — ending
  // with the narrator realizing why it matters and acting on it. Sent to
  // Telegram. Recent narrative "angles" are remembered so stories don't
  // repeat, even when the same habit comes up again.
  STORY: {
    // Master on/off switch for the meal-time stories.
    ENABLED: true,
    // Hours (0–23, spreadsheet timezone) to send a story: breakfast,
    // lunch, dinner. Re-run Install triggers after changing these.
    MEAL_HOURS: [8, 13, 19],
    // DocumentProperties key tracking which meal slots were sent today.
    // JSON: { date:'YYYYMMDD', hours:[8,13] }
    LAST_SENT_PROP: 'lastMealStory',
    // DocumentProperties key holding the recently-used angles (JSON array)
    // so we can ask the model to avoid repeating them.
    RECENT_ANGLES_PROP: 'storyAngles',
    // How many recent angles to remember/avoid.
    ANGLE_HISTORY: 8,
    // Narrative framings rotated to keep stories fresh. One is picked per
    // story (preferring ones not used recently) and fed to the prompt.
    ANGLES: [
      "a historical figure's untold quiet moment",
      'a gentle animal fable',
      'a near-future science-fiction vignette',
      'an ordinary morning that turns luminous',
      'an athlete on the edge of giving up',
      'a letter from your older, wiser self',
      "a grandparent's hard-won lesson",
      'a traveler alone in a foreign city',
      'a small failure that became a turning point',
      'a mentor and an apprentice',
      'a myth or legend retold',
      'a childhood memory rediscovered years later',
      'two strangers on a long train ride',
      'a craftsperson perfecting one small thing',
    ],
  },

  // ==================== CALENDAR SYNC CONFIG ====================
  // One-way Google Calendar -> Weekly grid. Reads events for the current
  // week and drops their titles into the matching day/time slots.
  CALENDAR: {
    // Master on/off switch for calendar sync.
    ENABLED: true,
    // Empty string => primary/default calendar; otherwise a calendar name.
    CALENDAR_NAME: '',
    // Skip all-day events (they don't map to a 30-min time slot).
    SKIP_ALL_DAY: true,
    // Synced event titles are arbitrary strings, not Tasks-list entries.
    // When true, remove the task-cell dropdown on cells written by the
    // sync so those titles aren't flagged "invalid". Set false to keep
    // the dropdown validation on every cell.
    CLEAR_DROPDOWN: true,
    // DocumentProperties key listing "row,col" cells written by the last
    // sync, so re-syncing can clear stale ones instead of duplicating.
    SYNCED_CELLS_PROP: 'calSyncCells',
  },

  // ==================== XP / LEVELS / BADGES CONFIG ====================
  XP: {
    // Level curve: XP needed to advance from level L to L+1 is
    // BASE + (L-1) * STEP. Cumulative thresholds: L1=0, L2=50, L3=125,
    // L4=225, L5=350, ... XP is lifetime and never reset by New Week.
    BASE: 50,
    STEP: 25,
    PROP_KEY: 'xpState', // JSON { xp:number, level:number }
    BADGES_PROP: 'xpBadges', // JSON array of earned badge ids
    QUEST_COUNT_PROP: 'questCompletions', // lifetime quest items completed
    // Badge thresholds.
    CENTURION_XP: 100,
    RISING_STAR_LEVEL: 5,
    WEEK_WARRIOR_STREAK: 7,
    QUEST_MASTER_COUNT: 10,
    EARLY_BIRD_HOUR: 9, // a positive score logged before this hour
  },

  // ==================== WEEKLY BOSS CONFIG ====================
  // One rotating weekly challenge. Deterministically chosen per week from
  // DEFS. Defeating it (progress >= target) grants REWARD_XP + a badge,
  // once per week. Progress for 'points' bosses is this week's Summary
  // total; 'quests' bosses count quest items completed this week.
  BOSS: {
    REWARD_XP: 20,
    PROP_KEY: 'weeklyBoss', // JSON { weekStart, bossId, defeated, questCount }
    DEFS: [
      { id: 'points_40', type: 'points', target: 40, emoji: '🐉', name: 'Point Dragon' },
      { id: 'points_60', type: 'points', target: 60, emoji: '👹', name: 'Score Ogre' },
      { id: 'quests_5', type: 'quests', target: 5, emoji: '🧙', name: 'Quest Warden' },
      { id: 'points_30', type: 'points', target: 30, emoji: '🦇', name: 'Focus Bat' },
    ],
  },

  // ==================== COLORS ====================
  COLORS: {
    TODAY_HIGHLIGHT: '#FFFF00', // Yellow
    POSITIVE: '#70AD47', // Green
    NEGATIVE: '#ED7D31', // Orange-Red
    NEUTRAL: '#FFC000', // Yellow/amber
    CURRENT_TIME: '#FFFF00', // Yellow for current hour
    CLEAR: '#FFFFFF',
    BUTTON_FILL: '#DCE6F1', // Light blue control-bar buttons
    QUEST_HIGHLIGHT: '#FFD700', // Gold — today's featured quest items
  },

  // ==================== DRIVE / ARCHIVE ====================
  DRIVE_ARCHIVE_FOLDER: 'Weekly Plan Archives',
};
