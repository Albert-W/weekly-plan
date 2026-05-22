# Changelog

All notable changes to the Weekly Plan add-in. Newest first.

## 2026-05-21 — 2026-05-22 — Refactor wave

A two-day cleanup pass that took the add-in from "works in production but
fragile to change" to "boring to extend". 17 commits, 22 of 23 backlog
tasks done, plus one new feature and a full test harness.

### TL;DR metrics

| Dimension                 | Before              | After                     |
|---------------------------|---------------------|---------------------------|
| Source files              | 8                   | **12** (single-concern)   |
| Total source lines        | 2,415               | 2,769                     |
| Largest single file       | `weekly.js` 1,211   | `weekly.js` **788** (−35%) |
| Tests                     | 0                   | **81** (~3s)              |
| Test files                | 0                   | 14                        |
| Office.js round-trips per habit click | up to ~18  | **≤ 3**                   |
| Office.js round-trips per score entry | 5–6        | **2**                     |
| Office.js round-trips per "New Week"  | ~231       | **2**                     |
| Inline `<script>` blocks in HTML | 1 (date icon) | **0**                  |
| Inline `onclick=` handlers       | 14            | **0**                  |
| Magic cell addresses outside CONFIG | 14         | **0**                  |
| `try/catch` boilerplate sites    | 12            | **5** (rest via `withStatus`) |
| Hardcoded sheet names in code    | 3             | **0** (Goals, Charter, others) |
| Dead code                        | 1 fn (`downloadXLSX`) | **0**          |

### What landed (in commit order)

| # | Commit  | Task | One-liner |
|---|---------|------|-----------|
| 1 | `5fbca1b` | #13 | Fix duplicate `onChanged` listener (was firing N× per sheet switch) |
| 2 | `332dd57` | #1  | Batch reads in `recordHabitDone` (N+1 → 2 syncs) |
| 3 | `bf29d3d` | #2  | Batch reads in `processWeeklyScoreChange` (5–6 → 2 syncs) |
| 4 | `e3397cd` | #15 | Batch reads in `clearForNewWeek` (~231 → 2 syncs) |
| 5 | `99a975a` | #14 | De-duplicate weekly CSV export (`buildWeeklyCSV` helper) |
| 6 | `06ad4a7` | —   | Vitest test suite + in-memory fake Excel context (50 tests) |
| 7 | `d218305` | #3  | De-duplicate habits refresh-dates |
| 8 | `e6340c6` | new | Background ticker auto-updates current-time highlight every minute |
| 9 | `cbf1b30` | #16 | Split task-creation out of `ui.js` into new `tasks.js` |
| 10 | `bfb69f6` | #4, #7, #8, #18 | Rename `CONFIG`/state fields and move magic strings into `CONFIG` |
| 11 | `6f68547` | #17 | Extract day-column math into `utils.js` helpers |
| 12 | `c1cfa05` | #11, #12 | Split `weekly.js` (1,212 lines) into `summary.js` + `export.js` + bigger `tasks.js`; drop dead `downloadXLSX` |
| 13 | `206887f` | #19 | Sheet-handler registry; `events.js` becomes pure routing |
| 14 | `22189cf` | #6  | `withStatus(label, fn)` wrapper; remove 7 try/catch boilerplate sites |
| 15 | `0d528f9` | #5, #21, #22, #23 | HTML cleanup + platform guards (`isExcelOnline`, Desktop CSV fallback) |
| 16 | `06bf196` | #10 | Centralize hardcoded cell addresses into `CONFIG` |
| 17 | `a22d686` | #9  | Compute weekly grid extent dynamically from the sheet itself |

### Per-file size before / after

```
                                Before  After  Δ
src/taskpane/js/app.js             239    290  +51   (button delegation, bootstrap, time-highlight ticker)
src/taskpane/js/config.js           79     97  +18   (new TASKS section, address constants)
src/taskpane/js/events.js          188    169  −19   (slimmed to pure routing)
src/taskpane/js/habits.js          275    266   −9   (de-duped refresh, batched recordHabitDone)
src/taskpane/js/state.js            32     38   +6   (changeHandler, grid extent defaults)
src/taskpane/js/ui.js              284    278   −6   (slimmed addTask, withStatus wrapper)
src/taskpane/js/utils.js           107    167  +60   (day-column helpers, isExcelOnline)
src/taskpane/js/weekly.js        1,211    788 −423   (split, batched, dedup'd)
src/taskpane/js/export.js          new    422   ↑    (CSV/archive helpers, extracted)
src/taskpane/js/summary.js         new    140   ↑    (Summary-sheet domain, extracted)
src/taskpane/js/tasks.js           new     76   ↑    (Tasks-sheet domain)
src/taskpane/js/registry.js        new     38   ↑    (per-sheet handler registry)
                                ─────  ─────
total                            2,415  2,769  +354  (across 12 files instead of 8)
```

The line *total* grew (+15%) because each new file has its own header
comment + `window.foo = foo` export block, but the **distribution of
complexity** improved dramatically: the largest single file shrank by
423 lines, and every file now has one clear concern.

### Highlights

#### Performance (Office.js round-trips)

The most user-visible win. Office.js bundles operations until you call
`context.sync()`; each sync is an IPC round-trip (~50–100 ms on Excel
Online). Code that ran a `sync()` inside a loop produced laggy clicks.

- **`recordHabitDone`** — 18 syncs (13-day streak) → **2 syncs**, regardless
  of streak length.
- **`processWeeklyScoreChange`** — 5–6 syncs → **2 syncs**, regardless of
  Tasks-sheet size.
- **`clearForNewWeek`** — ~231 syncs (7 days × 33 rows + overhead) →
  **2 syncs**. "New Week" went from a multi-second UI freeze to instant.

The test suite pins all three with `expect(getSyncCount()).toBe(2)` so a
regression breaks the build.

#### Bugs found and fixed along the way

- **Duplicate event handlers** — `registerOnChangedEvent` was called on
  init, refresh, and every sheet activation but never removed the
  previous handler. The score-change processor was firing 2×, 3×, … N×
  per click, silently inflating daily totals and the Summary sheet.
  Fixed in commit `5fbca1b`.
- **`others` row created with missing stats** — when an unknown task
  triggered the auto-created "others" fallback row, columns D / F / G
  were left null. Subsequent scores on that task did `null + score = NaN`
  in some browsers. Fixed in commit `bf29d3d`.
- **No background time-highlight refresh** — the yellow highlight on the
  current-time row only moved on user action. If you left the pane open
  from 14:30 to 15:00, the highlight stayed on the old row. Fixed in
  commit `e6340c6` (new feature, aligned 60s ticker).

#### Architecture

- **Single-concern files.** `weekly.js` was 1,212 lines doing five things.
  Now: weekly.js (weekly-sheet domain), `tasks.js`, `summary.js`,
  `export.js`, `registry.js` — each file is the answer to "where is
  X?".
- **Sheet-handler registry** (`registry.js`). `events.js` no longer knows
  any sheet names. Adding a new sheet means writing a domain file and
  calling `registerSheetHandlers(name, { onSelection, onActivate, onChange })`.
  Zero edits to events.js.
- **Sheet-driven grid extent.** `LAST_TIME_ROW`/`SCORE_ROW` used to be
  hardcoded as `36`/`38`. Now `initializeWeeklySheet` detects them from
  column B's used range and writes to `state.weekly.{lastTimeRow,scoreRow}`.
  CONFIG values became defaults. Users can resize the schedule grid in
  the spreadsheet — no code change required.

#### Cleanliness

- **`CONFIG` is the source of truth.** Every magic string (sheet names,
  cell addresses, the special `'others'` task name, the day-column
  arithmetic) moved into CONFIG or a derived helper. The codebase now
  contains **zero** bare cell-address literals outside `config.js`.
- **Consistent error UX.** `withStatus(label, fn)` (`ui.js`) wraps every
  user-triggered action and produces uniform `"{Action} failed: {msg}"`
  banners. Removed 7 copies of the same try/catch boilerplate.
- **HTML has no JavaScript.** The single inline `<script>` and all 14
  `onclick=` handlers moved to a `bootstrapDom()` function in app.js
  driven by `data-action` / `data-toggle` attributes.

#### Cross-platform robustness

- **`isExcelOnline()` helper** never throws even on older Office clients
  where `Office.PlatformType` is undefined.
- **CSV download** detects Excel Desktop (where `<a download>` is
  silently blocked by the sandboxed webview), copies the CSV to the
  clipboard, and shows a clear "paste into a new file" message instead
  of silently failing.

### Test infrastructure (new)

```
tests/
  setup.js                 Loads src/ into globals (no source changes needed)
  harness.js               Re-exports for ESM test files
  mocks/
    office.js              Minimal Office stub
    excel.js               In-memory fake of Excel object model with
                           queue/sync semantics and a sync counter
  utils.test.js            (15) date / column / address / platform helpers
  csv-helpers.test.js      (10) escapeCSV, formatExcelTime
  buildWeeklyCSV.test.js   (7)  header / time / filename / perf
  recordHabitDone.test.js  (7)  streak / counts / perf
  processWeeklyScoreChange.test.js (8) task lookup, others row, colors, perf
  clearForNewWeek.test.js  (5)  contract + exact-2-syncs perf guard
  events.test.js           (2)  handler de-duplication
  registry.test.js         (4)  sheet-handler routing
  createTask.test.js       (6)  task append + validation
  refreshHabitsDates.test.js (2) date window + state update
  timeHighlightTicker.test.js (2) routing skip-when-not-Weekly
  withStatus.test.js       (4)  success/error/no-banner-on-success
  platformGuards.test.js   (5)  isExcelOnline + downloadCSV Desktop fallback
  initializeWeeklySheet.test.js (4) grid extent detection
                          ────
                           81 tests, ~3s, runs via `npm test`
```

The fake Excel context (`tests/mocks/excel.js`, ~370 lines) was the
foundation. It implements `getRange`, `getUsedRange`, `getOffsetRange`,
`load('values'|'rowCount'|'rowIndex')`, the `values`/`format.fill.color`
setters, `clear()`, `onChanged.add/.remove`, `onSelectionChanged.add/.remove`,
and exposes a `syncCount` counter so tests can pin perf regressions
alongside correctness.

### What's still on the backlog

- **#20 — ES modules migration.** Convert from `window.foo = foo;`
  globals to `import`/`export`. Touches every file. Foundational rather
  than urgent — the project is no-build by design, and ESM would
  require either a build step or careful ordering of `<script type="module">`
  tags. Deferred.

### Method notes

- For sweep-style renames across many files (CONFIG/state field names,
  cell-address literals), `sed -i ''` across `find src tests -name '*.js'`
  was the right tool. BSD sed (macOS) doesn't honor `\b` word boundaries
  the way GNU sed does — when targets are unique enough (e.g.
  `state.weekly.taskl`) the literal prefix does the disambiguation.
- For the multi-file extraction in #11 (`weekly.js` split), a single
  throw-away Python script parsed function boundaries, copied blocks
  into new files, and rewrote the source. Safer than 15 sequential
  edit-tool calls on one giant file.
- The fake Excel context turned out to be the most-leveraged single
  investment: every domain test rides on it, and adding `rowIndex`
  support (for task #9) was a 4-line change.
- Most refactors landed alongside their tests. The few that didn't
  ride on existing tests as the safety net — "if 81 stay green, the
  rename was safe".

### Verification

`npm test` → 81 passed, 14 files, ~3 seconds, on Node 22 with Vitest
2.1 and jsdom 25.
