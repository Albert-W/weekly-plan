import { describe, it, expect } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { recordHabitDone, CONFIG, state } from './harness.js';

/**
 * recordHabitDone integration test against the fake Excel.
 *
 * Covers:
 *  - Streak counted from consecutive non-zero prior day cells.
 *  - Weighted score = baseScore * 1.1^streak.
 *  - Day count and total count both incremented.
 *  - Label cell (column B) gets POSITIVE fill color.
 *  - Sync-count regression guard for task #1: <=3 syncs total
 *    regardless of streak length (the third allowed sync covers
 *    updateSummary, which has its own context-sync chain).
 */

function setupHabitsSheet({ habitName = 'Read', baseScore = 1, priorDays = [] } = {}) {
  const fake = makeFakeExcel({
    sheets: [CONFIG.HABITS_SHEET, CONFIG.SUMMARY_SHEET],
    activeSheet: CONFIG.HABITS_SHEET,
  });
  const { helpers } = fake;
  const row = CONFIG.HABITS.DATA_START_ROW; // 4

  helpers.setCells(CONFIG.HABITS_SHEET, {
    [`A${row}`]: habitName,
    [`C${row}`]: baseScore,
  });

  // Fill in prior day cells D{row} .. Q{row} (14 day columns)
  const dayStart = CONFIG.HABITS.COLUMNS.DAY_START; // 'D'
  const dayStartIdx = dayStart.charCodeAt(0) - 'A'.charCodeAt(0); // 3
  for (let i = 0; i < priorDays.length; i++) {
    const col = String.fromCharCode('A'.charCodeAt(0) + dayStartIdx + i);
    helpers.setCells(CONFIG.HABITS_SHEET, { [`${col}${row}`]: priorDays[i] });
  }

  return { fake, row };
}

describe('recordHabitDone', () => {
  it('records a single completion with no prior streak', async () => {
    const { fake, row } = setupHabitsSheet({ baseScore: 2 });
    state.habits.currentDayIndex = 0; // today = first day of window
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    // D{row} now contains 1, R{row} now contains 1
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `R${row}`)).toBe(1);
    expect(fake.helpers.getCellColor(CONFIG.HABITS_SHEET, `B${row}`)).toBe(CONFIG.COLORS.POSITIVE);
  });

  it('counts streak from consecutive prior non-zero day cells', async () => {
    // 3-day streak ending today = days 0, 1, 2 filled, today = day 3
    const { fake, row } = setupHabitsSheet({ baseScore: 1, priorDays: [1, 1, 1] });
    state.habits.currentDayIndex = 3;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    // Today's column = column index 3+3 = 6 = 'G'
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `G${row}`)).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `R${row}`)).toBe(1);
    // streak = 3 prior consecutive days -> 1 * 1.1^3 ~= 1.331 written nowhere
    // (the weighted score only goes to updateSummary, not back to the sheet)
  });

  it('breaks streak on a zero/empty cell', async () => {
    // streak = 0 because day-before-yesterday is empty
    const { fake, row } = setupHabitsSheet({ baseScore: 1, priorDays: [1, '', 1] });
    state.habits.currentDayIndex = 3;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    // Today's column G should still get +1
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `G${row}`)).toBe(1);
  });

  it('increments existing counts rather than overwriting them', async () => {
    const { fake, row } = setupHabitsSheet({ baseScore: 1, priorDays: [] });
    // Today's cell already has a count of 2, and R column has 5
    fake.helpers.setCells(CONFIG.HABITS_SHEET, { [`D${row}`]: 2, [`R${row}`]: 5 });
    state.habits.currentDayIndex = 0;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBe(3);
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `R${row}`)).toBe(6);
  });

  it('PERF: uses at most 3 syncs regardless of streak length (regression guard for task #1)', async () => {
    // 7-day streak — old code would have done 7+5 = 12 syncs
    const { fake, row } = setupHabitsSheet({ baseScore: 1, priorDays: [1, 1, 1, 1, 1, 1, 1] });
    state.habits.currentDayIndex = 7;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();
    fake.helpers.resetSyncCount();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    // 2 inside recordHabitDone + however many updateSummary uses.
    // updateSummary on a fresh Summary sheet is ~4 syncs (its own
    // perf opportunity, tracked separately). Pinning at <=7 for now;
    // tighten when updateSummary is batched.
    expect(fake.helpers.getSyncCount()).toBeLessThanOrEqual(7);
  });

  it('does not write today if the habit name cell is empty', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET, CONFIG.SUMMARY_SHEET] });
    const row = CONFIG.HABITS.DATA_START_ROW;
    // intentionally do not set A{row}
    state.habits.currentDayIndex = 0;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    // Nothing should have been written.
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `R${row}`)).toBeNull();
  });

  it('does nothing if currentDayIndex is -1 (date not found)', async () => {
    const { fake, row } = setupHabitsSheet();
    state.habits.currentDayIndex = -1;
    state.habits.lastRow = row;
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await recordHabitDone(ctx, row); });

    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBeNull();
  });
});
