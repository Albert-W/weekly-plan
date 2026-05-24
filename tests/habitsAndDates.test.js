import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG } from './harness.js';

const findHabitsDayIndex = globalThis.findHabitsDayIndex;
const setNewWeekDates = globalThis.setNewWeekDates;

// Pin wall clock: Mon Jan 1 2024 15:30. Cross-midnight runs of
// tests below would otherwise see today.getDate() differ from
// the value seeded into days[3]. Task #38.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);

beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(FAKE_NOW); });
afterEach(() => { vi.useRealTimers(); });

/**
 * findHabitsDayIndex returns 0..13 if today's day-of-month matches a
 * value in D3:Q3, else -1.
 */
describe('findHabitsDayIndex', () => {
  it('finds today in the 14-day header window', async () => {
    const today = new Date().getDate();
    // Build 14 distinct "day numbers" guaranteed not to collide with today
    // except at index 3, where we place today exactly.
    const days = [];
    for (let i = 0; i < CONFIG.HABITS.DAYS_COUNT; i++) {
      // Use 50+i so values are out of the 1..31 range and can't accidentally
      // equal today's day-of-month.
      days.push(50 + i);
    }
    days[3] = today;

    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET] });
    const cells = {};
    'DEFGHIJKLMNOPQ'.split('').forEach((col, i) => { cells[`${col}3`] = days[i]; });
    fake.helpers.setCells(CONFIG.HABITS_SHEET, cells);
    fake.installAsExcelGlobal();

    let idx;
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);
      idx = await findHabitsDayIndex(ctx, sheet);
    });

    expect(idx).toBe(3);
  });

  it('returns -1 when today is not in the header', async () => {
    const today = new Date().getDate();
    const wrong = today === 1 ? 2 : 1;
    const days = new Array(CONFIG.HABITS.DAYS_COUNT).fill(wrong);

    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET] });
    const cells = {};
    'DEFGHIJKLMNOPQ'.split('').forEach((col, i) => { cells[`${col}3`] = days[i]; });
    fake.helpers.setCells(CONFIG.HABITS_SHEET, cells);
    fake.installAsExcelGlobal();

    let idx;
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.HABITS_SHEET);
      idx = await findHabitsDayIndex(ctx, sheet);
    });

    expect(idx).toBe(-1);
  });
});

/**
 * setNewWeekDates writes:
 *   B4 = "YYYY MM"
 *   D4, F4, H4, J4, L4, N4, P4 = day numbers Mon..Sun
 */
describe('setNewWeekDates', () => {
  it('writes B4 and 7 day-number headers', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await setNewWeekDates(ctx, sheet);
    });

    const b4 = fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'B4');
    expect(b4).toMatch(/^\d{4} \d{2}$/);

    // 7 day-number headers
    const cols = ['D', 'F', 'H', 'J', 'L', 'N', 'P'];
    const nums = cols.map((c) => fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `${c}4`));
    expect(nums.every((n) => typeof n === 'number' && n >= 1 && n <= 31)).toBe(true);

    // Consecutive days (allowing month wrap)
    for (let i = 1; i < 7; i++) {
      const diff = nums[i] - nums[i - 1];
      expect(diff === 1 || diff < 0).toBe(true);
    }
  });
});
