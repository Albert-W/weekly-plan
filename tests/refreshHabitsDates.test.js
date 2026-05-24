import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, getMonday } from './harness.js';

// Pin clock — refreshHabitsDates writes today's date number and a
// 14-day window starting at today's Monday. Cross-midnight runs would
// otherwise drift between the test's setup and assertion. Task #38.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);
beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(FAKE_NOW); });
afterEach(() => { vi.useRealTimers(); });

/**
 * refreshHabitsDates integration test.
 *
 * The public function refreshHabitsDates owns Excel.run and delegates
 * sheet-touching logic to refreshHabitsDatesWithContext (task #3
 * de-dup). This test exercises the public entry point so we catch
 * regressions in either layer.
 *
 * Behavior contract:
 *  - B3 set to "YYYY MM" of this week's Monday
 *  - D3:Q3 set to 14 consecutive day numbers starting from Monday
 *  - Data area D{start}:Q{lastRow} cleared
 *  - state.habits.currentDayIndex updated to today's slot in the
 *    new header
 */

describe('refreshHabitsDates', () => {
  it('writes the year/month and 14 day numbers, then updates currentDayIndex', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET], activeSheet: CONFIG.HABITS_SHEET });
    fake.installAsExcelGlobal();

    // Pretend a habit existed before the refresh
    fake.helpers.setCells(CONFIG.HABITS_SHEET, {
      A4: 'Read',
      D4: 1,
      E4: 1,
    });
    state.habits.lastRow = 4;

    await globalThis.refreshHabitsDates();

    // B3 = "YYYY MM" of this week's Monday
    const monday = getMonday(new Date());
    const expectedYearMonth = `${monday.getFullYear()} ${String(monday.getMonth() + 1).padStart(2, '0')}`;
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'B3')).toBe(expectedYearMonth);

    // D3 = Monday's day number
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'D3')).toBe(monday.getDate());
    // E3 = Monday + 1 day
    const tue = new Date(monday); tue.setDate(monday.getDate() + 1);
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'E3')).toBe(tue.getDate());
    // Q3 = Monday + 13 days (end of 14-day window)
    const last = new Date(monday); last.setDate(monday.getDate() + 13);
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'Q3')).toBe(last.getDate());

    // Old habit data area cleared
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'D4')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'E4')).toBeNull();
    // Habit name in column A is NOT touched (data area starts at D)
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'A4')).toBe('Read');

    // currentDayIndex now reflects today within the new header
    const today = new Date();
    const expectedIndex = Math.floor(
      (Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) -
        Date.UTC(monday.getFullYear(), monday.getMonth(), monday.getDate())) /
        86400000
    );
    expect(state.habits.currentDayIndex).toBe(expectedIndex);
  });

  it('skips clearing when there are no prior habit rows', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET], activeSheet: CONFIG.HABITS_SHEET });
    fake.installAsExcelGlobal();
    state.habits.lastRow = 3; // below DATA_START_ROW

    await expect(globalThis.refreshHabitsDates()).resolves.not.toThrow();
    // B3 still set
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'B3')).toBeTypeOf('string');
  });
});
