# How to Use Weekly Plan Add-in on Excel Web

This is a step-by-step guide to deploy and use your Weekly Plan add-in on Excel Online.

## Overview

```
┌─────────────────────────────────────────────────────────────┐
│  LOCAL DEVELOPMENT          │  PRODUCTION DEPLOYMENT        │
│  ─────────────────          │  ────────────────────         │
│  1. Run local HTTPS server  │  1. Push to GitHub            │
│  2. Load manifest-simple    │  2. Enable GitHub Pages       │
│  3. Edit → Save → Refresh   │  3. Update manifest URL       │
│     (instant reload!)       │  4. Upload to Excel Online    │
└─────────────────────────────────────────────────────────────┘
```

---

## Part A: Local Development (Recommended)

This is the best way to develop and test changes on the fly.

### Step 1: Generate SSL Certificates (One Time)

```bash
npx office-addin-dev-certs install
```

This creates certificates at `~/.office-addin-dev-certs/`

### Step 2: Start the Development Server

```bash
cd /Users/yichangwu/Documents/excel/weekly-plan/src/taskpane

npx http-server -c-1 -p 3000 --cors -S \
  -C ~/.office-addin-dev-certs/localhost.crt \
  -K ~/.office-addin-dev-certs/localhost.key
```

**Flags explained:**
- `-c-1` = No caching (changes load instantly!)
- `-p 3000` = Port 3000
- `--cors` = Allow cross-origin requests
- `-S` = HTTPS mode (required for add-ins)

### Step 3: Load Add-in in Excel Online

1. Go to [office.com](https://www.office.com) → Open **Excel**
2. Open your workbook (or create new)
3. **Insert** → **Add-ins** → **Upload My Add-in**
4. Select `manifest-simple.xml` from your project folder
5. Click **Upload**

### Step 4: Develop!

Now you can:
1. **Edit any JS file** in your editor
2. **Save** the file
3. **Refresh** the browser (F5)
4. **See changes immediately!** ✨

No server restart needed!

---

## Part B: Production Deployment (GitHub Pages)

For permanent hosting so others can use your add-in.

### Step 1: Create GitHub Repository

1. Go to [github.com](https://github.com) → Sign in
2. Click **"New repository"**
3. Name it: `weekly-plan`
4. Make it **Public** (required for free GitHub Pages)
5. Click **"Create repository"**

### Step 2: Upload Files

```bash
cd /Users/yichangwu/Documents/excel/weekly-plan

git init
git add .
git commit -m "Initial commit - Weekly Plan add-in"

git remote add origin https://github.com/YOUR_USERNAME/weekly-plan.git
git branch -M main
git push -u origin main
```

### Step 3: Enable GitHub Pages

1. Go to your repo on GitHub
2. **Settings** → **Pages** (left sidebar)
3. Under "Source", select **"Deploy from a branch"**
4. Choose **"main"** branch and **"/ (root)"** folder
5. Click **Save**
6. Wait 1-2 minutes

### Step 4: Update Manifest URL

Your add-in URL will be:
```
https://YOUR_USERNAME.github.io/weekly-plan/src/taskpane/taskpane.html
```

Edit `manifest.xml` and update the SourceLocation:
```xml
<SourceLocation DefaultValue="https://YOUR_USERNAME.github.io/weekly-plan/src/taskpane/taskpane.html"/>
```

Push the updated manifest:
```bash
git add manifest.xml
git commit -m "Update manifest URL"
git push
```

### Step 5: Load in Excel Online

1. Go to [office.com](https://www.office.com) → Excel
2. Open your workbook
3. **Insert** → **Add-ins** → **Upload My Add-in**
4. Upload `manifest.xml` (the production one)

---

## Project Structure

```
weekly-plan/
├── manifest-simple.xml      ← For local development (localhost:3000)
├── manifest.xml             ← For production (GitHub Pages URL)
├── README.md
├── TUTORIAL.md
└── src/
    └── taskpane/
        ├── taskpane.html    ← Main UI
        └── js/
            ├── config.js    ← All configuration values
            ├── state.js     ← Global state
            ├── utils.js     ← Helper functions
            ├── ui.js        ← UI updates & modals
            ├── habits.js    ← Habits sheet logic
            ├── weekly.js    ← Weekly sheet logic
            ├── events.js    ← Event handlers
            └── app.js       ← Main initialization
```

---

## Setting Up Your Workbook

### Required Sheets

| Sheet Name | Purpose |
|------------|---------|
| **Weekly** | Main timetable (auto-selected on open) |
| **Habits** | Daily habit tracking |
| **Tasks** | Task list with weights |
| **Summary** | Daily score aggregation |

### Weekly Sheet Structure

| Row | A | B | C | D | E | F | ... |
|-----|---|---|---|---|---|---|-----|
| 2 | Controls | | Help | | Add | | |
| 4 | | 2024 12 | Mon | 16 | Tue | 17 | ... |
| 5 | | 8:00 | Task | Score | Task | Score | ... |
| 6 | | 9:00 | Task | Score | Task | Score | ... |
| ... | | ... | | | | | |
| 36 | | 17:00 | Task | Score | Task | Score | ... |
| 38 | | Totals | | (sum) | | (sum) | ... |

### Habits Sheet Structure

| Row | A | B | C | D-Q (Days) | R |
|-----|---|---|---|------------|---|
| 3 | Habit | 2024 12 | Score | 16 17 18... | Total |
| 4 | Exercise | Health | 5 | | 0 |
| 5 | Reading | Learning | 3 | | 0 |

### Tasks Sheet Structure

| A (Name) | B (Weight) | C (Created) | D (Last Used) | E | F (Count) | G (Score) |
|----------|-----------|-------------|---------------|---|-----------|-----------|
| Deep Work | 2 | 20241218 | 20241218 | | 5 | 8.4 |
| Exercise | 1.5 | 20241218 | 20241218 | | 3 | 4.5 |

### Summary Sheet Structure

| A (Date) | B | C | D (Positive) | E (Negative) | F (Total) |
|----------|---|---|--------------|--------------|-----------|
| 20241218 | | | 5.6 | -0.4 | 5.2 |
| 20241217 | | | 4.2 | 0 | 4.2 |

---

## Using the Add-in

### On Open

When you open the file:
1. ✅ Weekly sheet is automatically selected
2. ✅ Current day column is highlighted (yellow)
3. ✅ Current time row is highlighted
4. ✅ If new week detected → previous week archived as CSV

### Weekly Timetable

| Action | How |
|--------|-----|
| Select a task | Click task cell → Choose from dropdown |
| Enter score | Click score cell → Type 0, 0.2, 0.4, 0.6, 0.8, or 1 |
| Random fill | Click **🎲 Random Pick** in sidebar |
| Archive week | Click **📦 Archive** → Downloads CSV |
| Start new week | Click **🗓️ New Week** → Clears data |
| Export summary | Click **📊 Export Summary** |

### Habits Tracker

| Action | How |
|--------|-----|
| Record habit | Click habit name in Column A |
| Sort by score | Click **📊 Sort** in sidebar |
| Refresh dates | Click **📅 Dates** in sidebar |

---

## Configuration

Edit `src/taskpane/js/config.js` to customize:

```javascript
const CONFIG = {
  // Sheet names (must match your workbook)
  WEEKLY_SHEET: 'Weekly',
  HABITS_SHEET: 'Habits',
  TASKS_SHEET: 'Tasks',
  SUMMARY_SHEET: 'Summary',

  // Weekly sheet layout
  WEEKLY: {
    DATA_START_ROW: 5,      // First data row
    lastTimeLine: 36,       // Last row with times
    scoreLine: 38,          // Row for daily totals
  },

  // Summary sheet columns
  SUMMARY: {
    DATE_COLUMN: 'A',
    POSITIVE_SCORE_COLUMN: 'D',
    NEGATIVE_SCORE_COLUMN: 'E',
    TOTAL_SCORE_COLUMN: 'F'
  }
};
```

---

## Troubleshooting

### "Invalid manifest" error
- Ensure manifest uses `https://` URL
- Check SourceLocation path is correct
- For local dev, server must be running on port 3000

### Add-in not loading latest changes
- Check server is running with `-c-1` flag (no cache)
- Try hard refresh: Cmd+Shift+R (Mac) / Ctrl+Shift+R (Windows)
- Remove and re-add the add-in

### "alert is not supported" error
- Fixed: We use `showWarningPopup()` instead of `alert()`
- Uses custom modal dialogs that work in Office Add-ins

### Selection events not firing
- Make sure you're on the correct sheet (Weekly or Habits)
- Check browser console for errors (F12)
- Try clicking **🔄 Refresh** in the sidebar

### Score not updating Summary
- Verify Summary sheet exists
- Check config has correct column letters (D, E, F)
- Look for errors in browser console

---

## Quick Reference

### Local Development Commands

```bash
# Generate SSL certs (one time)
npx office-addin-dev-certs install

# Start server (no caching)
cd src/taskpane
npx http-server -c-1 -p 3000 --cors -S \
  -C ~/.office-addin-dev-certs/localhost.crt \
  -K ~/.office-addin-dev-certs/localhost.key
```

### Manifest Files

| File | Use For | URL |
|------|---------|-----|
| `manifest-simple.xml` | Local dev | `https://localhost:3000/taskpane.html` |
| `manifest.xml` | Production | `https://YOUR_USERNAME.github.io/...` |

### Key Files

| File | What It Does |
|------|--------------|
| `config.js` | All settings in one place |
| `weekly.js` | Weekly sheet + archive + scores |
| `habits.js` | Habit tracking + streaks |
| `app.js` | Initialization, auto-select Weekly |
| `ui.js` | Modals, status messages |

---

## Summary

| Step | Local Dev | Production |
|------|-----------|------------|
| 1 | Generate SSL certs | Push to GitHub |
| 2 | Start HTTPS server | Enable GitHub Pages |
| 3 | Upload manifest-simple.xml | Update manifest.xml URL |
| 4 | Develop → Save → Refresh | Upload manifest.xml |

**Local dev time:** ~2 minutes
**Production setup:** ~10 minutes
**Cost:** Free! 🎉
