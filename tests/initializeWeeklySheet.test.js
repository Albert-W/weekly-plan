import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state } from './harness.js';

/**
 * initializeWeeklySheet contract tests.
 *
 * Auto-detect of lastTimeRow/scoreRow from the sheet was REMOVED by
 * user request. CONFIG (LAST_TIME_ROW=36, SCORE_ROW=38) is the
 * single source of truth. The function now only sets currentDayIndex
 * and lastMonday — grid extent stays at CONFIG defaults regardless
 * of what the sheet looks like.
 */

const initializeWeeklySheet = globalThis.initializeWeeklySheet;

describe('initializeWeeklySheet - CONFIG-driven (no auto-detect)', () => {
  beforeEach(() => {
    state.weekly.lastTimeRow = CONFIG.WEEKLY.LAST_TIME_ROW;
    state.weekly.scoreRow = CONFIG.WEEKLY.SCORE_ROW;
  });

  it('does NOT change lastTimeRow / scoreRow regardless of sheet content', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    // Seed the sheet with a wide range — auto-detect would have
    // bumped lastTimeRow to 50 here. We assert it stays at 36.
    for (let row = CONFIG.WEEKLY.DATA_START_ROW; row <= 50; row++) {
      fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { [`B${row}`]: 8 });
    }

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.lastTimeRow).toBe(CONFIG.WEEKLY.LAST_TIME_ROW); // 36
    expect(state.weekly.scoreRow).toBe(CONFIG.WEEKLY.SCORE_ROW);        // 38
  });

  it('sets currentDayIndex to a valid 0..6 value', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.currentDayIndex).toBeGreaterThanOrEqual(0);
    expect(state.weekly.currentDayIndex).toBeLessThan(7);
  });

  it('sets lastMonday to a Date', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.lastMonday).toBeInstanceOf(Date);
  });
});
