# Weekly Plan - Excel Add-in

This is a **unified Excel Add-in** for **Weekly Planning and Timetable** management with integrated **Habits Tracking**.

It uses **Office Add-ins with Office.js** which supports `SelectionChanged` events - the closest equivalent to VBA's `Worksheet_SelectionChange`.

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
