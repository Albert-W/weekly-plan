import { describe, it, expect, beforeEach } from 'vitest';
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
 */

const highlightCurrentTimeRow = globalThis.highlightCurrentTimeRow;

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
    state.weekly.currentDayIndex = 0;
    state.weekly.lastTimeRow = 36;
  });

  it('highlights the matching row when times are stored as whole-hour numbers', async () => {
    // 8, 9, 10, ..., 22 — pick the one matching the current hour
    const fake = setupTimeColumn([8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]);
    fake.installAsExcelGlobal();

    const hour = new Date().getHours();
    // Determine expected row
    let expectedRow = -1;
    for (let i = 0; i < 15; i++) {
      if (8 + i <= hour) expectedRow = CONFIG.WEEKLY.DATA_START_ROW + i;
    }

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    if (expectedRow > 0) {
      expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${expectedRow}`))
        .toBe(CONFIG.COLORS.CURRENT_TIME);
    }
  });

  it('highlights the matching row when times are stored as fraction-of-day', async () => {
    // 8/24, 9/24, ..., 22/24 — same expected behavior
    const fractions = [];
    for (let h = 8; h <= 22; h++) fractions.push(h / 24);
    const fake = setupTimeColumn(fractions);
    fake.installAsExcelGlobal();

    const hour = new Date().getHours();
    let expectedRow = -1;
    for (let i = 0; i < fractions.length; i++) {
      if (8 + i <= hour) expectedRow = CONFIG.WEEKLY.DATA_START_ROW + i;
    }

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    if (expectedRow > 0) {
      expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${expectedRow}`))
        .toBe(CONFIG.COLORS.CURRENT_TIME);
    }
  });

  it('also highlights the current-day task + score cells when they are empty', async () => {
    // All-day grid; force a known match by seeding only the row we want
    const fake = setupTimeColumn([8]); // single row at 8am
    state.weekly.lastTimeRow = CONFIG.WEEKLY.DATA_START_ROW;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentTimeRow(ctx, sheet);
    });

    const row = CONFIG.WEEKLY.DATA_START_ROW;
    // Highlight only fires if current hour >= 8. Run defensively.
    if (new Date().getHours() >= 8) {
      expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `B${row}`))
        .toBe(CONFIG.COLORS.CURRENT_TIME);
      // Monday columns C/D are highlighted because empty
      expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `C${row}`))
        .toBe(CONFIG.COLORS.CURRENT_TIME);
    }
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

    // B always highlighted on match; C and D should NOT be re-highlighted
    if (new Date().getHours() >= 8) {
      // C/D still have whatever color was set when the score was entered (null
      // since we only seeded the value). The contract: NOT CURRENT_TIME.
      expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, `C${row}`))
        .not.toBe(CONFIG.COLORS.CURRENT_TIME);
    }
  });
});
