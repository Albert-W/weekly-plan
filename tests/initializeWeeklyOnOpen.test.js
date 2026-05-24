import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import {
  CONFIG, state, formatDateYYYYMMDD, getMonday,
  initializeWeeklyOnOpen,
} from './harness.js';

// Pin wall clock — Mon Jan 1 2024 15:30. Cross-midnight / cross-week
// runs would otherwise see formatDateYYYYMMDD(today) drift between
// the test's setup and assertion. Task #38.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);

/**
 * initializeWeeklyOnOpen is the per-open orchestrator. It:
 *  - Reads B4 ("YYYY MM") and D4 (first day-of-month).
 *  - Computes diffDays = days since that Monday.
 *  - If diffDays >= 7 -> archive the previous week (calls
 *    archiveWeekAutomatically and downloadCSV), clear the data area,
 *    set new dates.
 *  - Otherwise: leaves data alone.
 *  - Always recomputes currentDayIndex + lastMonday, highlights
 *    current day + current time, and stamps lastInitDate.
 */

function setupOpenedSheet({ b4, d4 }) {
  const fake = makeFakeExcel({
    sheets: [CONFIG.WEEKLY_SHEET, CONFIG.SUMMARY_SHEET],
  });
  fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B4: b4, D4: d4 });
  // Seed a time slot so highlightCurrentTimeRow has something to bind to.
  fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B5: 8 });
  // Mimic Office.context platform = Online so downloadCSV would actually
  // try to write a blob; the dirty stub avoids needing a real URL API.
  if (typeof URL.createObjectURL !== 'function') {
    URL.createObjectURL = () => 'blob:fake';
    URL.revokeObjectURL = () => {};
  }
  globalThis.Office.context.platform = 'OfficeOnline';
  return fake;
}

describe('initializeWeeklyOnOpen', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(FAKE_NOW);
    state.weekly.lastInitDate = null;
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('first-time use: no valid date in B4 -> sets new-week dates', async () => {
    const fake = setupOpenedSheet({ b4: '', d4: '' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeWeeklyOnOpen(ctx); });

    // B4 now has "YYYY MM" of the current Monday
    const b4 = fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'B4');
    expect(b4).toMatch(/^\d{4} \d{2}$/);

    // currentDayIndex is set
    expect(state.weekly.currentDayIndex).toBeGreaterThanOrEqual(0);
    expect(state.weekly.currentDayIndex).toBeLessThan(7);
    // lastMonday is set
    expect(state.weekly.lastMonday).toBeInstanceOf(Date);
    // lastInitDate stamped to today
    expect(state.weekly.lastInitDate).toBe(formatDateYYYYMMDD(new Date()));
  });

  it('same week (diffDays < 7) leaves B4 alone and updates state only', async () => {
    // Pretend we're 2 days past this week's Monday.
    const monday = getMonday(new Date());
    const yearMonth = `${monday.getFullYear()} ${String(monday.getMonth() + 1).padStart(2, '0')}`;
    const fake = setupOpenedSheet({ b4: yearMonth, d4: monday.getDate() });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeWeeklyOnOpen(ctx); });

    // B4 unchanged
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'B4')).toBe(yearMonth);
    // state populated
    expect(state.weekly.lastInitDate).toBe(formatDateYYYYMMDD(new Date()));
  });

  it('new week detected (>=7 days since recorded Monday) -> overwrites B4 with current week', async () => {
    // Recorded Monday is 2 weeks ago.
    const today = new Date();
    const old = new Date(today);
    old.setDate(today.getDate() - 14);
    const oldMonday = getMonday(old);
    const oldYearMonth = `${oldMonday.getFullYear()} ${String(oldMonday.getMonth() + 1).padStart(2, '0')}`;

    const fake = setupOpenedSheet({ b4: oldYearMonth, d4: oldMonday.getDate() });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeWeeklyOnOpen(ctx); });

    // B4 should now be the CURRENT week's Monday, not the old one.
    const currentMonday = getMonday(today);
    const expectedYearMonth = `${currentMonday.getFullYear()} ${String(currentMonday.getMonth() + 1).padStart(2, '0')}`;
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'B4')).toBe(expectedYearMonth);
  });
});
