import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, formatDateYYYYMMDD, updateSummary } from './harness.js';

// Pin clock — updateSummary writes today's row and the assertions
// expect that row. Cross-midnight runs would otherwise see them
// disagree. Task #38.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);
beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(FAKE_NOW); });
afterEach(() => { vi.useRealTimers(); });

/**
 * updateSummary tests.
 *
 * Contract:
 *  - If a row for today's YYYYMMDD already exists in column A,
 *    accumulate into it.
 *  - Otherwise append a new row and bump state.weekly.lastSummaryRow.
 *  - positiveScore > 0 accumulates into Summary col D.
 *  - negativeScore < 0 accumulates into Summary col E.
 *  - Total in col F = (col D) + (col E).
 *  - No-op silently if the Summary sheet is missing.
 */

function setup() {
  const fake = makeFakeExcel({ sheets: [CONFIG.SUMMARY_SHEET] });
  fake.installAsExcelGlobal();
  state.weekly.lastSummaryRow = 0;
  return fake;
}

describe('updateSummary', () => {
  beforeEach(() => { state.weekly.lastSummaryRow = 0; });

  it('creates a new row when today is not yet recorded', async () => {
    const fake = setup();
    await Excel.run(async (ctx) => { await updateSummary(ctx, 1.5, 0); });

    const today = formatDateYYYYMMDD(new Date());
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'A1')).toBe(today);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'D1')).toBe(1.5);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'F1')).toBe(1.5);
    expect(state.weekly.lastSummaryRow).toBe(1);
  });

  it('accumulates into the existing row for today', async () => {
    const fake = setup();
    await Excel.run(async (ctx) => { await updateSummary(ctx, 1.0, 0); });
    await Excel.run(async (ctx) => { await updateSummary(ctx, 0.5, 0); });

    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'D1')).toBeCloseTo(1.5, 6);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'F1')).toBeCloseTo(1.5, 6);
    expect(state.weekly.lastSummaryRow).toBe(1);
  });

  it('routes negative scores into column E and reflects them in the total', async () => {
    const fake = setup();
    await Excel.run(async (ctx) => { await updateSummary(ctx, 0, -0.4); });

    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'D1')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'E1')).toBe(-0.4);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'F1')).toBe(-0.4);
  });

  it('handles a positive + negative pair in a single call', async () => {
    const fake = setup();
    await Excel.run(async (ctx) => { await updateSummary(ctx, 1.0, -0.3); });

    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'D1')).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'E1')).toBe(-0.3);
    expect(fake.helpers.getCellValue(CONFIG.SUMMARY_SHEET, 'F1')).toBeCloseTo(0.7, 6);
  });

  it('is a no-op when the Summary sheet does not exist', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    // Should not throw.
    await expect(
      Excel.run(async (ctx) => { await updateSummary(ctx, 1, 0); })
    ).resolves.not.toThrow();
  });
});
