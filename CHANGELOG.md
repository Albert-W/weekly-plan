# Changelog

All notable changes to the Weekly Plan add-in. Newest first.

## 2026-06-18 — Meal-time AI habit stories (Google Sheets edition)

Added a second Telegram cron: three times a day a random habit is picked and
**Google Gemini** writes a short, original motivating story about it.

- **New `google-apps-script/Story.js`** — picks a random habit
  (`pickRandomHabit_`), chooses a fresh narrative angle (rotated + recency-aware
  so stories never repeat), prompts Gemini via `generateContent`
  (`generateStory_`), and sends the result to Telegram as plain text
  (`sendMealStory_`). Per-meal-slot dedup + enable gate; manual preview via
  `sendMealStoryNowFromUI`. Gemini key setup via `setUpGemini`.
- **3 new `mealStory` time triggers** at `CONFIG.STORY.MEAL_HOURS`
  (default 8:00 / 13:00 / 19:00), registered by `installTriggers`.
- **`Telegram.js`** — `sendTelegramMessage_` now takes an optional `parseMode`
  (default `HTML`; `null` sends plain text, used for LLM prose).
- **Config** — added `CONFIG.GEMINI` (key prop, model, temperature, token cap)
  and `CONFIG.STORY` (enabled, meal hours, dedup + angle-history props, angle
  list).
- **Menu** — added **Send meal story now** and **Set up Gemini…**.
- **Docs** — README_GAS.md gains a "Meal-time habit stories" section + file-map
  entry. (Reuses the `script.external_request` scope added for Telegram.)

## 2026-06-18 — Telegram morning notification (Google Sheets edition)

Added a daily **Telegram** morning recap as the primary notification channel,
replacing the email recap (which is now disabled by default).

- **New `google-apps-script/Telegram.js`** — builds and sends the morning recap
  (yesterday's score, today's Daily Quest habit + task, combo, XP/level, weekly
  boss) via the Telegram Bot API (`sendMessage`, HTML parse mode) using
  `UrlFetchApp`. Enable-gated + deduped once per day.
- **Credentials in Script Properties**, never in source — set once via the new
  **Weekly Plan → Set up Telegram…** menu item (`setUpTelegram`), which prompts
  for the bot token + chat id and sends a confirmation message. Preview any time
  with **Send Telegram recap now**.
- **New `morningTelegram` time trigger** (~`CONFIG.TELEGRAM.SEND_HOUR`, default
  8am) registered by `installTriggers`. `dailyMaintenance` (~5am) no longer sends
  the morning notification.
- **Config** — added `CONFIG.TELEGRAM` (ENABLED, BOT_TOKEN_PROP, CHAT_ID_PROP,
  SEND_HOUR, LAST_SENT_PROP); set `CONFIG.EMAIL.ENABLED = false`.
- **Manifest** — added the `script.external_request` OAuth scope for outbound
  HTTPS to the Telegram API.
- **Docs** — README_GAS.md gains a "Morning Telegram recap" setup section and
  file-map entry.

> Note: this feature lives in the Apps Script edition only; the `src/` Office
> add-in (and its Vitest suite) is unchanged.

## 2026-05-23 — 2026-05-24 — Test expansion, UI cleanup, correctness fire-drill

Three smaller waves rolled together after the initial refactor: more test
coverage, a UI audit-and-cleanup pass, and a post-refactor code review
that turned up four real bugs the first pass missed.

### TL;DR metrics (delta from previous changelog entry)

| Dimension                 | Then       | Now                        |
|---------------------------|------------|----------------------------|
| Source files              | 12         | **13** (+ `concurrency.js`) |
| Source lines              | 2,769      | **2,997** (+228)            |
| Test files                | 14         | **27** (+13)                |
| Tests                     | 81         | **153** (+72)               |
| Test lines                | ~1,323     | **2,679** (+1,356)          |
| Above-the-fold pane chrome | ~150 px   | **~60 px** (−60%)          |
| Today's score visibility  | hidden (row 38) | **always on top of Weekly card** |
| Race vulnerability (RMW)  | open       | **closed + regression test pinned** |
| CSV formula injection     | exploitable | **guarded + 8 prefix tests** |
| `showModal` HTML injection | open      | **textContent + ARIA role=dialog** |
| Sheet-driven grid extent  | never executed at runtime | **wired into `initializeWeeklyOnOpen`** |

### Commits (newest first)

| Commit  | What | Wave |
|---------|------|------|
| `dbb05bd` | **[H4] Serialize per-sheet handlers** to fix RMW races (queue approach) | Correctness |
| `e83d689` | **[M11]** `columnLetterToIndex` fix for multi-letter columns; drop "broken" test pin | Correctness |
| `909f274` | **[H3]** `showModal`: replace `innerHTML` with safe DOM construction (XSS defense) | Correctness |
| `f5d5d56` | **[H2]** `escapeCSV`: guard against CSV formula injection (OWASP) | Correctness |
| `e4d6746` | **[H1]** Wire `initializeWeeklySheet` into `initializeWeeklyOnOpen` — task #9 is no longer dead | Correctness |
| `6611464` | Footer: replace inaccurate "Works on Desktop, Web & Mobile" with honest status | UI |
| `9e444ba` | Add **Today's Score widget** to the task pane | UI |
| `1d78f8d` | Tighten task pane layout: inline status, merge sheet into title | UI |
| `4c604f0` | Remove dead UI: Save button, Delete panel, Thank panel | UI |
| `ef4288f` | Add 13 more tests covering init orchestrators and DOM-form integration | Tests |
| `6eccf3e` | Add 27 more tests across 6 new files covering previously-untested behavior | Tests |

### Wave 1 — Test expansion (+40 tests)

Audited every `src/` function against existing tests, identified the
highest-leverage gaps, and added focused integration tests for:

- `randomPick` — fills only empty slots, preserves user tasks, no-op on rows without time labels, routes per `currentDayIndex`, no-op on empty Tasks list.
- `highlightCurrentTimeRow` — whole-hour values, fraction-of-day values, score-cell highlighting only when empty.
- `findHabitsDayIndex` + `setNewWeekDates` — date lookups and header writes.
- `updateSummary` — new-row creation, accumulation, +/- routing, missing-sheet no-op.
- `exportSheetAsCSV` — serialization, time formatting, comma/quote escaping, missing-sheet handling.
- `sheetSelection` — Habits column-A click, Weekly task/score click contracts (3 scenarios each).
- `initializeHabitsSheet` + `highlightCurrentDay` — auto-refresh when today missing, header highlight cleanup.
- `refreshTimeHighlight` — silent flag, missing-sheet branch.
- `addTask` — DOM form integration (happy path, empty name, bad weight).
- `initializeWeeklyOnOpen` — first-time, same-week, new-week detection.

While writing these, surfaced and fixed a real fake-Excel gap: `getOffsetRange`/`getUsedRange` returned `Range` objects that bypassed `attachValuesSetter()`, so subsequent writes silently didn't queue. The fake now treats both paths identically.

### Wave 2 — UI cleanup (5 commits, 1 new feature)

After a UI audit identified dead elements, layout waste, and a missing core feature:

- **Removed dead UI.** Save button (its toast duplicated the static note below it). Delete button + panel (just told you to use the Tasks sheet manually — not a feature). Thank panel (orphan: no in-pane button opened it). Plus the corresponding `CONFIG.WEEKLY.BUTTONS.DELETE`/`THANK` entries and their sheet-side handlers.
- **Tightened the layout.** Inline status banner (was a whole card). Merged "Current Sheet" indicator into the H1 title ("Weekly Plan · Habits"). Dropped "Common Actions" card wrapper. ~90 vertical pixels reclaimed above the fold.
- **NEW: Today's Score widget.** Three-cell strip (`+ POSITIVE  |  TODAY  |  − NEGATIVE`) at the top of the Weekly card. Surfaces what the whole add-in is *for*. Refreshes on: init, sheet activation, every score change, and the background time-highlight ticker (free — already running every minute). Backed by a new read-only `getTodayScore(context)` in `summary.js`.
- **Honest footer.** Replaced "Works on Desktop, Web & Mobile" with "Best on Excel Online • Desktop supported (CSV via clipboard)" — accurate to the post-#22 reality (Desktop CSV uses clipboard fallback; mobile is untested).

### Wave 3 — Correctness fire-drill (the post-refactor audit)

A fresh code-review audit flagged 15 items, including **four high-severity bugs** the original refactor missed. All four are now closed and pinned by regression tests.

#### [H1] `initializeWeeklySheet` was never called from production

The CHANGELOG bragged about "Sheet-driven grid extent" — that `state.weekly.lastTimeRow/scoreRow/currentDayIndex/lastMonday` were detected from the actual sheet at init. **They weren't.** The detector was defined and tested but never invoked. In real Excel, those fields kept their CONFIG defaults forever. Any user who extended the time grid past row 36 hit silent truncation in five different code paths.

Fix: one `await initializeWeeklySheet(context)` at the top of `initializeWeeklyOnOpen`, plus removal of a dead `B:B` getUsedRange block that loaded the value and threw it away. Net +1 useful line, ~10 dead lines deleted. New regression test seeds B5..B50 and asserts `lastTimeRow === 50`.

#### [H2] CSV formula injection

`escapeCSV` handled `, " \n` but not `= + - @ \t \r`. A task literally named `=HYPERLINK("http://evil","go")` would execute when the user opened the exported CSV — classic OWASP CSV-injection.

Fix: prefix string-typed values starting with risky chars with a single quote. The `typeof value === 'string'` gate is critical — without it, numeric scores like `-0.4` would be turned into the literal text `'-0.4` and break downstream analysis. 8 new tests pin both the guard and the numeric-preservation contract.

#### [H3] `showModal` interpolated into `innerHTML`

`modal.innerHTML = \`... ${title} ... ${message} ...\`` — caller-controlled strings going straight into HTML. All current callers passed static strings, but the API was XSS-by-design.

Fix: rebuild with `createElement` + `textContent` for the title and message slots. As a bonus, added `role="dialog"`, `aria-modal="true"`, and `aria-labelledby` (partial credit toward the a11y task). 8 new tests including the canonical `<img src=x onerror=alert(1)>` payload-as-text assertion.

#### [H4] Read-modify-write race in score/habit handlers

Three handlers (`recordHabitDone`, `processWeeklyScoreChange`, `updateSummary`) all do load-then-add-then-write. Office.js fires `onChanged` per cell, and may dispatch handlers concurrently when the user pastes or types fast. Two handlers can each read `X`, each write `X + delta`, losing one increment. Tests ran handlers serially so this never surfaced.

Fix: new `src/taskpane/js/concurrency.js` exposing `serializeSheetWrite(sheetName, fn)`. Per-sheet `Map` of Promise chains — same sheet serializes, different sheets parallelize. `.catch(() => {})` keeps the chain alive after a thrown handler. Wrapped in `events.js#handleSelectionChanged` and `#handleCellChanged` at the outer level so the entire handler lifecycle is serialized. The three RMW function bodies are untouched.

**Strongest evidence the test is real:** temporarily replaced `serializeSheetWrite`'s body with `return fn();` and re-ran the suite. Both new race regression tests **failed**, as expected. Restored.

#### Pattern bonus: [M11] `columnLetterToIndex` for multi-letter columns

`columnLetterToIndex('AA')` returned 0 (same as 'A'), because the math treated A=0 instead of base-26-with-no-zero. The bug was pinned by a test with the comment "currently broken". Today's grid stops at column Q so the bug didn't bite — but task #9 introduces a real risk of growth past Z.

Fix: standard base-26-with-no-zero algorithm. Removed the "broken on purpose" comment, replaced with an exhaustive **round-trip property test** that asserts `columnLetterToIndex(indexToColumnLetter(i)) === i` for all 702 single+double-letter columns.

### Lesson: a test pinning a known bug is worse than no test

Three of the Wave-3 items (H1, M11) were *originally noted* in earlier work but ratified rather than fixed:
- H1 had a comment `// Using fixed values from CONFIG instead of dynamic calculation`.
- M11 had a test comment `// NOTE: columnLetterToIndex is currently broken for multi-letter`.

A test that knowingly asserts a buggy behavior makes the bug *harder* to fix — the next person reverts your fix because the test "regressed." The cleanup pattern: fix the production code, replace the pin with a proper assertion, and add an exhaustive property test (round-trips, parameterized inputs) when possible.

### Architectural patterns introduced

- **`serializeSheetWrite(sheetName, fn)`** — per-resource async write queue, swallowing rejections to keep the chain alive. Reusable for any future RMW pattern.
- **Wrap at the boundary, not the body.** The race fix lives at the dispatch site (`events.js`), so handler functions stay ignorant of concurrency. Same shape as `withStatus`: the cross-cutting concern lives in the wrapper, not the wrapped.
- **Property tests for invertible functions.** `columnLetterToIndex` now has an exhaustive sweep instead of a handful of point assertions.

### What's still on the post-audit backlog

Wave-A correctness items are all closed. Remaining backlog (filed as tasks #33–#43) is all pattern/quality polish — none risk silent data loss or security exploits:

- Removing 5 dead exports.
- Replacing `'C5:Z...'` magic in `clearForNewWeek` with derived CONFIG.
- Replacing the `app.js` try/catch wall with a `safeInit` helper.
- Awaiting `refreshTodayScoreWidget` everywhere (no more fire-and-forget).
- Batching `updateSummary` to 2 syncs + tightening the placeholder `<=7` perf guards.
- Fake timers in time-of-day-dependent tests.
- README refresh against the 13-file actual layout.
- Full a11y attribute pass.
- Misc orphan-JSDoc and duplicate-helper cleanup.

### Verification

`npm test` → **153/153 pass**, 27 test files, ~5 seconds on Node 22 with Vitest 2.1 + jsdom 25.

---

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
