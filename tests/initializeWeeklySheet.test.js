import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state } from './harness.js';

/**
 * initializeWeeklySheet detection tests (task #9).
 *
 * The function now derives state.weekly.lastTimeRow and scoreRow
 * from the actual sheet (column B's used range) instead of using
 * the hardcoded CONFIG values. Falls back to CONFIG defaults when
 * the sheet is empty.
 */

const initializeWeeklySheet = globalThis.initializeWeeklySheet;

function seedTimeColumn(fake, lastRow) {
  // Put a time label in B5..B{lastRow}. The function reads B:B's
  // used range, so we need at least one cell populated.
  const cells = {};
  for (let row = CONFIG.WEEKLY.DATA_START_ROW; row <= lastRow; row++) {
    cells[`B${row}`] = `${row - CONFIG.WEEKLY.DATA_START_ROW + 8}:00`;
  }
  fake.helpers.setCells(CONFIG.WEEKLY_SHEET, cells);
}

describe('initializeWeeklySheet - sheet-driven grid extent', () => {
  beforeEach(() => {
    // Reset to CONFIG defaults before each test.
    state.weekly.lastTimeRow = CONFIG.WEEKLY.LAST_TIME_ROW;
    state.weekly.scoreRow = CONFIG.WEEKLY.SCORE_ROW;
  });

  it('detects the default 36-row grid', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    seedTimeColumn(fake, 36);

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.lastTimeRow).toBe(36);
    expect(state.weekly.scoreRow).toBe(38);
  });

  it('detects an EXTENDED grid (user added time slots)', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    seedTimeColumn(fake, 50);

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.lastTimeRow).toBe(50);
    expect(state.weekly.scoreRow).toBe(52);
  });

  it('detects a SHRUNK grid (user removed time slots)', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    seedTimeColumn(fake, 20);

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    expect(state.weekly.lastTimeRow).toBe(20);
    expect(state.weekly.scoreRow).toBe(22);
  });

  it('falls back to CONFIG defaults when column B is empty', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    // No cells set — used range is empty.
    state.weekly.lastTimeRow = 999;  // sentinel to detect overwrite
    state.weekly.scoreRow = 999;

    await Excel.run(async (ctx) => { await initializeWeeklySheet(ctx); });

    // Sentinel preserved because detection found nothing useful
    // (detectedLastTimeRow < DATA_START_ROW).
    expect(state.weekly.lastTimeRow).toBe(999);
    expect(state.weekly.scoreRow).toBe(999);
  });
});
