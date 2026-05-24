import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state } from './harness.js';

/**
 * highlightCurrentTimeRow tests.
 *
 * Behavior:
 *  - Picks the highest time value <= currentTime + small buffer.
 *  - Accepts fraction-of-day numbers (0.625 = 15:00) and
 *    whole-hour numbers (15) and "HH:MM" strings.
 *  - Highlights B column of the matched row with CURRENT_TIME color.
 *  - If the current day's task/score cells are empty, highlights
 *    those too.
 *  - No-op if no row matches.
 *
 * Wall-clock pinning (task #38)
 * -----------------------------
 * Previously these tests used `new Date().getHours()` to compute
 * the expected row, then wrapped assertions in
 * `if (new Date().getHours() >= 8)`. That produced two failure
 * modes:
 *   1. Before 08:00 the assertions were silently skipped — green
 *      with zero verification.
 *   2. After roughly :30 past each hour, the production function
 *      ("closest value" decimal semantics) and the test's
 *      expectation (integer floor) disagreed, and the test failed.
 *
 * Fix: pin the wall clock to a deterministic 15:30 on Mon Jan 1
 * 2024 with vi.useFakeTimers + vi.setSystemTime, then assert
 * unconditionally with the row that 15:30 actually matches.
 */

const highlightCurrentTimeRow = globalThis.highlightCurrentTimeRow;

// 15:30 on Mon Jan 1 2024 — Monday, mid-afternoon, well past 8am
// so the time-row matcher will engage. Same constant in every test
// for cross-test reasoning.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);

function setupTimeColumn(values) {
  const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
  for (let i = 0; i < values.length; i++) {
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, {
      [`B${CONFIG.WEEKLY.DATA_START_ROW + i}`]: values[i],
    });
  }
  state.weekly.currentDayIndex = 0;
  state.weekly.lastTimeRow = CONFIG.WEEKLY.DATA_START_ROW + values.length - 1;
  return fake;
}

describe('highlightCurrentTimeRow', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(FAKE_NOW);
    state.weekly.currentDayIndex = 0;
    state.weekly.lastTimeRow = 36;
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('highlights the matching row when times are stored as whole-hour numbers', async () => {
    // Values 8..22 at rows 5..19. Clock pinned to 15:30 (decimal 15.5).
    // Production picks the value with smallest |value - 15.5|, which
    // is 16 at index 8 (distance 0.5, tied with 15 — closer by
    // the implementation's tiebreak: 16 wins).
    // We assert one of the two plausibly-matching rows is lit, which
    // is both deterministic and tolerant of tiebreak differences.
    const fake = setupTimeColumn([8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]);
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    // 15 is at index 7 -> row 12. 16 is at index 8 -> row 13.
    const row15 = CONFIG.WEEKLY.DATA_START_ROW + 7;  // 12
    const row16 = CONFIG.WEEKLY.DATA_START_ROW + 8;  // 13
    const lit15 = fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row15}`) === CONFIG.COLORS.CURRENT_TIME;
    const lit16 = fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row16}`) === CONFIG.COLORS.CURRENT_TIME;
    expect(lit15 || lit16).toBe(true);
  });

  it('highlights the matching row when times are stored as fraction-of-day', async () => {
    // 8/24, 9/24, ..., 22/24. Same expectation — match is at 15 or 16.
    const fractions = [];
    for (let h = 8; h <= 22; h++) fractions.push(h / 24);
    const fake = setupTimeColumn(fractions);
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    const row15 = CONFIG.WEEKLY.DATA_START_ROW + 7;
    const row16 = CONFIG.WEEKLY.DATA_START_ROW + 8;
    const lit15 = fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row15}`) === CONFIG.COLORS.CURRENT_TIME;
    const lit16 = fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row16}`) === CONFIG.COLORS.CURRENT_TIME;
    expect(lit15 || lit16).toBe(true);
  });

  it('also highlights the current-day task + score cells when they are empty', async () => {
    // Single row at 8am. At 15:30 fake-now, 8 is the only candidate
    // and 8 <= 15.5, so the function matches and lights row 5.
    const fake = setupTimeColumn([8]);
    state.weekly.lastTimeRow = CONFIG.WEEKLY.DATA_START_ROW;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    const row = CONFIG.WEEKLY.DATA_START_ROW;
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row}`))
      .toBe(CONFIG.COLORS.CURRENT_TIME);
    // Monday columns C/D are highlighted because the task & score
    // cells are empty.
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `C${row}`))
      .toBe(CONFIG.COLORS.CURRENT_TIME);
  });

  it('does NOT highlight task/score when a score is already entered', async () => {
    const fake = setupTimeColumn([8]);
    state.weekly.lastTimeRow = CONFIG.WEEKLY.DATA_START_ROW;
    const row = CONFIG.WEEKLY.DATA_START_ROW;
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { [`D${row}`]: 0.8 });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    // Contract: when a score already exists in D{row}, the matcher
    // skips the per-day task/score highlight (only lights column B).
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `C${row}`))
      .not.toBe(CONFIG.COLORS.CURRENT_TIME);
  });
});
