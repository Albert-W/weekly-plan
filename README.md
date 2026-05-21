# Weekly Plan - Excel Add-in

## Introduction

**Weekly Plan** is a personal productivity Excel Add-in that turns a spreadsheet into a lightweight **weekly time-block planner** and **habits tracker**. It was migrated from a legacy VBA macro workbook to a modern, cross-platform **Office Add-in built on Office.js**, so the same workflow now runs in Excel Online, Excel for Mac, and Excel for Windows — no macros required.

The add-in lives in a side **task pane** next to your spreadsheet and reacts to what you do in the grid:

- **Weekly / Timetable sheet** — Plan each day in fixed time blocks (rows = hours, columns = days). Pair each block with a task from your `Tasks` sheet, then score how the block actually went (`0` – `1`). Scores are color-coded (green / yellow / red), summed into a daily total row, and rolled up into the `Summary` sheet automatically. A "Random Pick" button fills empty slots from your weighted task list, and a "New Week" action archives the current week to CSV and resets the grid.
- **Habits sheet** — Track up to 14 days of habit completions in a rolling window. Clicking a habit name records a completion and applies a streak bonus (`score × 1.1^streak`), rewarding consistency. Habits can be re-sorted by score from the task pane.
- **Auto-initialization** — When the workbook opens, the add-in activates the `Weekly` sheet, highlights the current day column and the current time row, and detects whether a new week (or a new day) has started so it can archive and reset for you.

### Why an Office Add-in instead of VBA or Google Sheets?

The original workbook relied heavily on VBA events (`Workbook_Open`, `Worksheet_SelectionChange`, `Worksheet_Change`) that don't exist in Google Sheets and don't run in Excel Online. Office.js is the only option that preserves the event-driven UX while also working in the browser and on multiple platforms.

| Feature | VBA | Google Sheets | Excel Add-in (this project) |
|---------|-----|---------------|------------------------------|
| Selection Change Event | ✅ | ❌ | ✅ |
| Cell Change Event | ✅ | ✅ | ✅ |
| Works Online | ❌ | ✅ | ✅ |
| Works on Desktop | ✅ | ❌ | ✅ |
| Works on Mobile | ❌ | ✅ | ✅ (limited) |
| Language | VBA | JavaScript | JavaScript |

### Architecture at a glance

The code is a plain HTML + vanilla-JS task pane (no build step, no framework) split into small modules so each Office.js concern stays isolated:

- `app.js` — `Office.onReady` bootstrap and per-sheet initialization
- `config.js` — All sheet names, row/column layout, colors, and scoring options
- `state.js` — Lightweight global state (current sheet, last-init date, counters)
- `weekly.js` / `habits.js` — Per-sheet business logic (scoring, archive, streaks)
- `events.js` — `onSelectionChanged` and `onChanged` handlers
- `ui.js` / `utils.js` — Status messages, modal popups, date and column helpers

Because there is no bundler, you can edit a JS file, hit refresh in Excel, and see changes instantly — see [Quick Start](#quick-start---local-development) below.

It uses **Office Add-ins with Office.js**, whose `SelectionChanged` event is the closest equivalent to VBA's `Worksheet_SelectionChange` and is what made this migration possible.

## Why Excel Add-ins?

| Feature | VBA | Google Sheets | Excel Add-in |
|---------|-----|---------------|--------------|
| Selection Change Event | ✅ | ❌ | ✅ |
| Cell Change Event | ✅ | ✅ | ✅ |
| Works Online | ❌ | ✅ | ✅ |
| Works on Desktop | ✅ | ❌ | ✅ |
| Works on Mobile | ❌ | ✅ | ✅ (limited) |
| Language | VBA | JavaScript | JavaScript |

## Features

### Auto-initialization on Open
- **Weekly sheet auto-activates** when you open the file
- **Current day column** is highlighted
- **Current time row** is highlighted
- **New week detection** - automatically archives and clears data

### Weekly Timetable (Primary Feature)
- Plan your week with time blocks
- Random task fill for empty slots
- Score tracking with color coding (green/yellow/red)
- Daily totals in score line (row 38)
- Summary sheet updates automatically
- Archive week data as CSV
- Start new week with one click

### Habits Tracker
- Click habit name to record completion
- Streak bonus calculation (Score × 1.1^streak)
- 14-day rolling window
- Sort habits by score

## Project Structure

```
weekly-plan/
├── manifest-simple.xml     # Simplified manifest for local dev
├── manifest.xml            # Full manifest for production
├── README.md               # This file
├── TUTORIAL.md             # Deployment guide
└── src/
    └── taskpane/
        ├── taskpane.html   # Main HTML UI
        └── js/
            ├── config.js   # Configuration constants
            ├── state.js    # Global state management
            ├── utils.js    # Utility functions
            ├── ui.js       # UI functions & modals
            ├── habits.js   # Habits sheet logic
            ├── weekly.js   # Weekly sheet logic
            ├── events.js   # Event handlers
            └── app.js      # Main initialization
```

## Configuration (config.js)

All configurable values are in one place:

```javascript
const CONFIG = {
  // Sheet names
  HABITS_SHEET: 'Habits',
  WEEKLY_SHEET: 'Weekly',
  TASKS_SHEET: 'Tasks',
  SUMMARY_SHEET: 'Summary',

  // Weekly sheet settings
  WEEKLY: {
    DATA_START_ROW: 5,
    lastTimeLine: 36,      // Last row with time data
    scoreLine: 38,         // Score totals row
    // ...
  },

  // Summary sheet columns
  SUMMARY: {
    DATE_COLUMN: 'A',
    POSITIVE_SCORE_COLUMN: 'D',
    NEGATIVE_SCORE_COLUMN: 'E',
    TOTAL_SCORE_COLUMN: 'F'
  },

  // Colors
  COLORS: {
    POSITIVE: '#70AD47',   // Green
    NEGATIVE: '#ED7D31',   // Red
    NEUTRAL: '#FFC000',    // Yellow
    // ...
  }
};
```

## Quick Start - Local Development

### 1. Start the development server (with no caching)

```bash
cd /Users/yichangwu/Documents/weekly-plan/src/taskpane
npx http-server -c-1 -p 3000 --cors -S \
  -C ~/.office-addin-dev-certs/localhost.crt \
  -K ~/.office-addin-dev-certs/localhost.key
```

### 2. Load the add-in in Excel Online

1. Go to [office.com](https://www.office.com) → Excel
2. Open your workbook
3. **Insert** → **Add-ins** → **Upload My Add-in**
4. Select `manifest-simple.xml`

### 3. Develop on the fly

- Edit any JS file → Save
- Refresh browser (F5)
- Changes appear instantly! ✨

## Key Features Explained

### Score Processing

When you enter a score in the Weekly sheet:

1. **Individual cell** - colored based on score
2. **Task cell** - colored to match
3. **Daily total (row 38)** - updated automatically
4. **Summary sheet** - columns D, E, F updated
5. **Tasks sheet** - count and total score updated

### Warning Popups

Office Add-ins don't support `alert()`. We use custom modals instead:

```javascript
showWarningPopup('Please select a task first!');
// Shows a styled modal in the taskpane
```

### Week Archive

At the start of a new week:
1. Previous week data exported as CSV
2. Weekly sheet cleared
3. New dates set automatically

Or manually:
- Click **📦 Archive** to export CSV
- Click **🗓️ New Week** to clear and reset

## Supported Sheets

### 1. Weekly/Timetable Sheet (Primary)
| Feature | VBA Function | Add-in Equivalent |
|---------|--------------|-------------------|
| Auto-select on open | `Workbook_Open` | `initializeAddin()` |
| Highlight current time | `hourTask()` | `highlightCurrentTimeRow()` |
| Random fill tasks | `RandomPick()` | `randomPick()` |
| Score tracking | `Worksheet_Change` | `processWeeklyScoreChange()` |
| Daily total update | Manual | Automatic (row 38) |
| Summary updates | Direct cell update | `updateSummary()` |

### 2. Habits Sheet
| Feature | VBA Function | Add-in Equivalent |
|---------|--------------|-------------------|
| Mark habit done | Double-click Column A | Click Column A |
| 14-day rolling window | `Worksheet_Activate` | Auto-refresh |
| Streak bonus (1.1^streak) | `Worksheet_BeforeDoubleClick` | `recordHabitDone()` |
| Sort by score | `ListSort()` | Sidebar button |

### 3. Supporting Sheets
- **Tasks** - Task list with weights (Column A: name, B: weight)
- **Summary** - Daily score aggregation (D: positive, E: negative, F: total)

## Limitations vs VBA

| VBA Feature | Office.js Support |
|-------------|-------------------|
| `ThisWorkbook.Save` | ❌ Not supported (use auto-save) |
| `Application.OnTime` | ❌ Not supported |
| `window.alert()` | ❌ Use custom modals |
| `SendKeys` | ❌ Not supported |
| Double-click event | ❌ Use single click |

## Files Overview

| File | Description |
|------|-------------|
| `config.js` | All configuration constants |
| `state.js` | Global state (current sheet, counters) |
| `utils.js` | Date formatting, column conversion |
| `ui.js` | Status messages, modals, UI updates |
| `habits.js` | Habits sheet logic |
| `weekly.js` | Weekly sheet + archive + summary |
| `events.js` | Selection/change event handlers |
| `app.js` | Main initialization, Office.onReady |

## Troubleshooting

### "Current sheet always Loading..."
- Clear browser cache (Cmd+Shift+R)
- Remove and re-add the add-in
- Check console for errors (F12)

### "alert is not supported"
- Fixed: We use `showWarningPopup()` with custom modals

### Changes not appearing
- Server running with `-c-1` flag?
- Try refreshing with F5
- Check if correct manifest is loaded

## License

Personal use. Migrated from VBA to Office.js.
