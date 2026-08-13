# Weekly Plan — Google Sheets edition (Apps Script)

This folder is a **Google Apps Script** port of the Office.js Excel add-in in
`../src`. It gives you the same weekly time-block planner + habits tracker
inside Google Sheets, running automatically whenever you open the sheet — plus
a few upgrades that the Office stack couldn't do (auto-scaffold, score
dropdowns + conditional-format colors, native checkboxes, real dialogs, an
in-sheet Archive tab, and a daily maintenance trigger).

## What you get

- **Auto-init on open** — a "Weekly Plan" menu + sidebar appear; the Weekly
  sheet activates, today's column and the current time row are highlighted, and
  a new week is detected → archived → reset.
- **Weekly timetable** — pick a task in a task column (C, E, G…), then choose a
  score (0–1) from the dropdown in the score column (D, F, H…). Score cells are
  colored by conditional formatting; the matching task cell is colored by the
  script; the daily total row, the Summary sheet, and Tasks stats all update.
- **Habits tracker** — tick a habit's **checkbox** to record a completion with a
  streak bonus (`base × 1.1^streak`). Sort by score / refresh the 14-day window
  from the sidebar.
- **Random pick / Add task** from the sidebar (Add Task is also on the menu as a
  prompt).
- **Week rollover** — archives the finished week into an **Archive** sheet *and*
  a CSV in a Google **Drive** folder, then clears the grid. Runs on open and via
  a **daily trigger** even when the sheet is closed. "New Week" asks for
  confirmation first.
- **Morning Telegram recap** — every morning (~8am) a Telegram bot sends
  yesterday's score, today's Daily Quest (featured habit + task), your combo,
  XP/level, and the weekly boss. Set it up once via **Weekly Plan → Set up
  Telegram…** (see below). The legacy email recap is disabled by default.

## File map (port of `../src/taskpane/js`)

| Apps Script file | Ported from | Purpose |
| --- | --- | --- |
| `Config.js` | `config.js` | All constants (sheet names, layout, colors, scoring) |
| `Utils.js` | `utils.js` | Date + base-26 column helpers, GAS sheet helpers |
| `Setup.js` | *(new)* | `setUpSheets()` auto-builds the whole structure |
| `Tasks.js` | `tasks.js` | Create/delete task, last-row lookup, sort by weight |
| `Summary.js` | `summary.js` | Daily score accumulation, today's score, CSV export |
| `Export.js` | `export.js` | CSV build, Archive-sheet append, Drive save |
| `Weekly.js` | `weekly.js` | Init, highlights, random pick, score processing, rollover |
| `Habits.js` | `habits.js` | Checkbox completion, streaks, sort, date window |
| `Quest.js` | *(new)* | Daily Quest: per-day habit+task pick, bonus, highlights, streak combos |
| `Email.js` | *(new)* | Morning summary email (yesterday's recap + today's quest) — disabled by default in favor of Telegram |
| `Telegram.js` | *(new)* | Morning recap via a Telegram bot (yesterday's score + today's quest, ~8am) |
| `Calendar.js` | *(new)* | One-way Google Calendar → Weekly grid import |
| `Xp.js` | *(new)* | XP, levels & badges (lifetime progression) |
| `Boss.js` | *(new)* | Weekly Boss: rotating objective, HP bar, XP+badge reward |
| `Triggers.js` | `app.js` + `events.js` | Menu, sidebar, onEdit dispatch, daily trigger, init |
| `Sidebar.html` | `taskpane.html` | Sidebar UI (`google.script.run` wiring) |
| `appsscript.json` | *(new)* | Manifest: V8, timezone, OAuth scopes |
| `.clasp.json` / `.claspignore` | *(new)* | clasp push config |

> The script files use the `.js` extension so the `clasp` CLI can push them
> (clasp stores them as `.gs` server-side). In Apps Script all files share one
> global scope, so the old `window.*` exports are unnecessary and were removed.

## Setup — option A: clasp push (recommended)

[clasp](https://github.com/google/clasp) pushes this folder straight into the
bound Apps Script project.

```bash
npm install -g @google/clasp     # one-time
clasp login                      # one-time, opens a browser

cd google-apps-script
```

Then point `.clasp.json` at your script:

- **Existing Sheet:** open the Sheet → **Extensions → Apps Script → Project
  Settings → Script ID**, copy it, and replace `<YOUR_SCRIPT_ID>` in
  `.clasp.json`. (Or run `clasp clone <SCRIPT_ID>` in an empty dir and copy the
  generated `.clasp.json` here.)
- **New standalone project:** `clasp create --type sheets --title "Weekly Plan"`
  (this creates a new Sheet + bound script and writes a `.clasp.json`).

Push and you're done:

```bash
clasp push
```

Then reload the Sheet, click **Set up sheets**, then **Install triggers** and
authorize (steps 7–9 below).

## Setup — option B: manual paste (~5 minutes)

1. Create (or open) the Google Sheet you want to use.
2. **Extensions → Apps Script**. This opens the bound script project.
3. For each `*.js` file in this folder: create a matching **Script** file in the
   editor and paste the contents.
   - `Config, Utils, Setup, Tasks, Summary, Export, Weekly, Habits, Quest, Email, Telegram, Calendar, Xp, Boss, Triggers`
4. Create an **HTML** file named `Sidebar` and paste `Sidebar.html` into it.
   (The editor names it `Sidebar.html` — do **not** include the `.html` in the
   name field.)
5. (Optional) Open **Project Settings → “Show appsscript.json manifest”**, then
   paste `appsscript.json` over the generated manifest so the timezone and OAuth
   scopes match.
6. **Save** the project.
7. Reload the Google Sheet. A **Weekly Plan** menu appears and the sidebar opens.
8. In the sidebar (or the menu): click **Set up sheets** to build the layout,
   then **Install triggers (run once)** and **authorize** when prompted (Sheets,
   Drive, UI, and trigger scopes).
9. Done — from now on, opening the sheet auto-runs init and opens the sidebar.

## Resulting sheet layout

- **Weekly** — clickable control bar in row 2 (**Help · Add Task · Delete Task ·
  Random Fill · Thanks**); `B4` = "yyyy mm"; day-name labels in row 4 task
  columns with day numbers (right-aligned) in the score columns; time blocks in
  `B5:B36` as decimal half-hours (**8 → 23.5**); task columns C/E/G/I/K/M/O
  with **dropdowns sourced from the Tasks list** (pick, don't type); score
  columns D/F/H/J/L/N/P with 0–1 dropdowns; calendar borders; daily totals in
  the **Scores** row (38).
- **Habits** — `A` = habit name, `B` = **Done? checkbox**, `C` = base score,
  `D:Q` = 14-day window (header in row 3), `R` = total count.
- **Tasks** — `A` name, `B` weight, `C` created, `D` last done, `F` count,
  `G` total score (headers in row 3, data from row 4).
- **Summary** — `A` date (YYYYMMDD), `D` positive, `E` negative, `F` total
  (header row 1).
- **Archive** — finished weeks appended with a `Week` label column.

CSV archives are saved to a Drive folder named **Weekly Plan Archives**.

## Morning Telegram recap

Each morning (~`CONFIG.TELEGRAM.SEND_HOUR`, default **8am**, spreadsheet
timezone) a Telegram bot DMs you the same recap the email used to send:
yesterday's score, today's Daily Quest (featured habit + task), combo, XP/level,
and the weekly boss. It's driven by the `morningTelegram` time trigger and
deduped once per day.

**One-time setup:**

1. In Telegram, message **@BotFather**, send `/newbot`, and follow the prompts to
   get a **bot token** (looks like `123456789:AA...`).
2. Send any message to your new bot (so it's allowed to DM you), then message
   **@userinfobot** to get your numeric **chat id**.
3. In the Sheet: **Weekly Plan → Set up Telegram…**, paste the token then the
   chat id. A test message is sent immediately so you can confirm it works.
4. Make sure triggers are installed (**Weekly Plan → Install triggers**) and
   authorize the new *external request* scope when prompted.

Use **Weekly Plan → Send Telegram recap now** any time to preview the message.
Credentials live in **Script Properties** (`telegramBotToken` / `telegramChatId`)
— never in source. Turn the feature off via `CONFIG.TELEGRAM.ENABLED`, or change
the delivery hour via `CONFIG.TELEGRAM.SEND_HOUR` (re-run Install triggers after
changing the hour). To switch back to the email recap, set
`CONFIG.EMAIL.ENABLED = true`.

## Notable differences vs. the Excel add-in (intentional)

- Habit completion via **checkbox**, not cell selection (no `onSelectionChange`).
- Score entry via **dropdown**; coloring via **conditional formatting** (survives
  sorts/recalc) instead of per-edit background writes.
- **Real dialogs**: Add Task prompt (menu) and a **New Week confirmation**.
- Rollover archives to an **Archive sheet + Drive CSV**, and also runs on a
  **daily time-driven trigger**.
- The live time-highlight + today-score refresh is driven by a **client-side
  timer in the sidebar** (Apps Script has no persistent server timer).
- Status messages use Sheets **toasts** + the sidebar status line.

## Troubleshooting

- **Menu/sidebar didn't appear** — reload the sheet; ensure `onOpen` saved
  without errors.
- **Edits don't update totals** — run **Install triggers** and authorize. The
  edit handler is an *installable* trigger (`handleEdit`). The row-2 control
  buttons use the built-in `onSelectionChange` *simple* trigger (fires
  automatically, no install needed).
- **Scores entered on mobile didn't tally** — the live trigger doesn't run in the
  Sheets mobile app. Back on desktop, use **Weekly Plan → Recalculate now** (or
  the sidebar's 🧮 button) to re-derive the daily Scores row + colors from the
  grid.
- **CSV export failed** — re-authorize; the Drive scope is `drive.file`, which
  lets the script manage only files/folders it creates.
- **"Today not found" on a habit** — click **Dates** to refresh the 14-day
  window.
