# CLAUDE.md

## Project Overview

**Weekly Plan** is a personal productivity and self-discipline gamification system built as a **Google Apps Script (GAS) add-on** bound to a Google Sheet. It's a port of an earlier Office.js Excel add-in (`../src`, not in this repo).

Core features:
- **Weekly timetable** — 7-day time-block grid (30-min slots, 08:00–23:30) with per-slot task selection and 0–1 scoring
- **Habits tracker** — 14-day rolling window with checkboxes; streak bonus (`base × 1.1^streak`)
- **Daily Quest** — one featured habit + task per day with bonus multipliers and consecutive-day combo streak
- **XP / Levels / Badges** — lifetime progression: Centurion, Rising Star, Week Warrior, Quest Master, Early Bird, Boss Slayer
- **Weekly Boss** — rotating weekly objective; defeat grants XP + badge
- **Daily Diary** — evening Google Form (link sent via Telegram ~21:30) → Diary sheet + "写日记" habit + Weekly top band; local SQLite + reader on the Mac (see `local/`)
- **Automation** — daily triggers: ~5am maintenance + calendar sync, ~8am Telegram recap, ~21:30 diary reminder
- **Web App API** (`WebApp.js`) — JSON snapshot + habit check-in + diary view (`?view=diary`) endpoints for an external Mac Mini bot

## Tech Stack

- **Language:** JavaScript (ES6+), Google Apps Script **V8** runtime
- **Host:** Google Sheets bound script (`scriptId: 1Vy1XPcREDvFI_90MCmGqrAbpq-MJZvb6I8FVYBdGkqaLMzYKQRgEft16`)
- **Timezone:** Europe/Dublin
- **Deployment:** `@google/clasp` CLI — `clasp push` to upload `.js`, `.html`, `appsscript.json`
- **No npm dependencies, no build step, no tests**
- **Google services:** SpreadsheetApp, CalendarApp, DriveApp, MailApp, UrlFetchApp, LockService, PropertiesService, HtmlService, ContentService, ScriptApp
- **External APIs:** Telegram Bot API (recaps + diary reminders)
- **OAuth scopes** (`appsscript.json`): `spreadsheets.currentonly`, `script.container.ui`, `script.scriptapp`, `script.send_mail`, `script.external_request`, `calendar.readonly`, `drive.file`, `forms`

## Architecture

All `.js` files share **one global scope** — no imports/exports. The `CONFIG` constant (defined in `Config.js`) is the single shared configuration object visible everywhere. Functions ending in `_` are conventionally "private."

### File Map

| File | Purpose |
|------|---------|
| `Config.js` | Global `CONFIG` object: sheet names, column layouts, scoring, colors, DocumentProperties key registry, feature toggles |
| `Utils.js` | Shared helpers: dates, base-26 math, spreadsheet access, activity log, toast, lock helpers |
| `Triggers.js` | Entry points: `onOpen` (menu + sidebar), `handleEdit` (installable onEdit), `onSelectionChange`, trigger installer, daily automation, sidebar wrappers (`*FromUI`) |
| `Weekly.js` | Timetable grid: init, week rollover, score processing (`processWeeklyScoreChange`), recalculation |
| `Habits.js` | Habits sheet: 14-day window refresh, checkbox completion (`recordHabitDone`), streak + quest combo |
| `Summary.js` | Daily score accumulation, today-score reads, CSV export |
| `Quest.js` | Daily Quest: FNV-1a deterministic daily pick, bonus multipliers, quest streak combo system |
| `Xp.js` | XP/levels/badges: level curve math, `awardXp_`, idempotent badge awarding |
| `Boss.js` | Weekly Boss: deterministic pick, progress tracking, `checkBossDefeat_` (XP + badge under lock) |
| `Email.js` | Morning summary email (HTML + text); disabled by default in favor of Telegram |
| `Telegram.js` | Telegram bot: credential setup, message send, morning recap builder |
| `Diary.js` | Daily diary: form setup (`setUpDiary`), onFormSubmit handler, Diary sheet upsert, Weekly top band, deferred habit/badge processing, WebApp diary payload |
| `Calendar.js` | One-way Google Calendar → Weekly grid import |
| `Export.js` | CSV building (with OWASP formula-injection guard), Archive sheet, Drive folder/file helpers |
| `Setup.js` | `setUpSheets()`: auto-scaffold/repair all sheets, dropdowns, conditional formatting, checkboxes |
| `WebApp.js` | Web App entry (`doGet`/`doPost`): daily snapshot JSON + habit check-in, auth via `syncAuthKey` |
| `Sidebar.html` | Client-side sidebar (320px), polls server via `google.script.run` every 3s/60s |

### Key Patterns

- **Locking:** `withLock_(fn)` / `tryWithLock_(fn)` in `Utils.js` — DocumentLock used for all scoring paths to prevent concurrent write conflicts
- **Deterministic picks:** Quest (daily) and Boss (weekly) are chosen via FNV-1a hash on `date|salt`, ensuring consistent results without stored state
- **Sidebar communication:** Server functions named `*FromUI` are called from `Sidebar.html` via `google.script.run`; polling every 3s for score/log/quest/xp/boss, every 60s for time-row highlight
- **Idempotent setup:** `setUpSheets()` and badge awarding check for existing state before acting — safe to re-run
- **Feature toggles:** Per-feature `ENABLED` flags on the config objects: `EMAIL.ENABLED`, `TELEGRAM.ENABLED`, `DIARY.ENABLED`, `CALENDAR.ENABLED`

### Cross-Module Data Flow (Scoring)

```
Triggers.handleEdit
  → Weekly.processWeeklyScoreChange   (timetable scores)
  → Habits.recordHabitDone            (habit checkboxes)
      ↓
  acquire document lock
  apply quest/combo multipliers (Quest.js)
  release lock
  flush sheet, updateSummary (Summary.js)
  awardXp_ + badge checks (Xp.js)
  checkBossDefeat_ (Boss.js)
  markQuestDone_ → recordQuestCompletion_ + bumpBossQuestCount_ (Quest.js)
```

### Secret / State Storage

- **Script Properties:** `telegramBotToken`, `telegramChatId`, `syncAuthKey`, `spreadsheetId`, `diaryFormId`, `diaryLastSent`, `diaryHabitLastDate`
- **DocumentProperties:** All runtime state under `wp.*`-prefixed keys (registry in `Config.js` lines 9–23) — except the diary's state, which deliberately lives in Script Properties because its onFormSubmit trigger fires in a cross-document (Form) context
- API keys never appear in source code

## Development Workflow

### Deploy Changes

```bash
clasp push   # uploads *.js, *.html, appsscript.json (per .claspignore)
```

### First-Time Setup (as documented in README_GAS.md)

1. `npm install -g @google/clasp && clasp login`
2. `clasp push`
3. In the Sheet: **Weekly Plan → Set up sheets → Install triggers** → authorize

### External Service Setup

- **Telegram:** **Weekly Plan → Set up Telegram** — paste bot token from @BotFather + chat ID from @userinfobot
- **Web App / Mac bot:** Deploy `WebApp.js` as Web App ("Execute as: me, Access: Anyone"), set `syncAuthKey` + `spreadsheetId` in Script Properties

### Trigger Schedule

- `dailyMaintenance` — ~5am (init + calendar sync + diary band refresh / deferred habit)
- `morningTelegram` — ~8am (recap message; includes last night's diary worry + plan)
- `diaryReminder` — ~21:30 (diary form link via Telegram; the `handleDiarySubmit` form trigger is installed separately by **Set up Diary…**)

All triggers are installed idempotently via `Triggers.installTriggers()`.

## Important Constraints

- **Single global scope** — no `import`/`export`; all symbols are shared. Avoid naming collisions across files.
- **No local run target** — code only executes in the Apps Script runtime. There is no `package.json`, no `npm start`, no test runner.
- **Office.js predecessor** (`../src`) is **not in this repo**. References in README/comments to `../src/taskpane/js/*` are historical.
- **Weekly grid row indices are hardcoded** (`DATA_START_ROW=5`, `SCORE_ROW=38`) — never insert rows above the grid. The diary top band uses the free rows 1 & 3 with wrap/height instead.
- **Form-trigger caveat**: `handleDiarySubmit` runs in a Form (cross-document) context — keep gamification side effects (habit/XP/badge) in spreadsheet-context functions (`processPendingDiaryHabits_`, `maybeAwardChronicler_`), and keep diary state in Script Properties.
- Sidebar is **poll-based** (Apps Script has no server→client push), keep polling intervals reasonable.
- All scoring mutations must go through the lock path to avoid concurrent-write corruption.
