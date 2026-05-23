import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, initializeHabitsSheet, highlightCurrentDay } from './harness.js';

/**
 * initializeHabitsSheet behavior:
 *  - Sets state.habits.lastRow from the used range
 *  - If today's day-of-month isn't in the header window, calls
 *    refreshHabitsDatesWithContext to repopulate it.
 *  - Highlights the current-date header cell with TODAY_HIGHLIGHT.
 */
describe('initializeHabitsSheet', () => {
  beforeEach(() => {
    state.habits.lastRow = 4;
    state.habits.currentDayIndex = -1;
  });

  it('auto-refreshes the date window when today is not found', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET] });
    // Seed header with dummy dates so getUsedRange isn't empty.
    fake.helpers.setCells(CONFIG.HABITS_SHEET, {
      D3: 99, E3: 99, F3: 99, G3: 99, H3: 99,
      I3: 99, J3: 99, K3: 99, L3: 99, M3: 99,
      N3: 99, O3: 99, P3: 99, Q3: 99,
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeHabitsSheet(ctx); });

    // After refresh, B3 should hold a "YYYY MM" string set by the refresh path.
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'B3')).toMatch(/^\d{4} \d{2}$/);
    // currentDayIndex should now point at a valid 0..13 slot for today.
    expect(state.habits.currentDayIndex).toBeGreaterThanOrEqual(0);
    expect(state.habits.currentDayIndex).toBeLessThan(CONFIG.HABITS.DAYS_COUNT);
  });

  it('highlights the current-date header cell after init', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await initializeHabitsSheet(ctx); });

    // Walk D3..Q3 and verify exactly one cell has the TODAY_HIGHLIGHT color.
    let highlighted = 0;
    'DEFGHIJKLMNOPQ'.split('').forEach((col) => {
      const c = fake.helpers.getCellColor(CONFIG.HABITS_SHEET, `${col}3`);
      if (c === CONFIG.COLORS.TODAY_HIGHLIGHT) highlighted++;
    });
    expect(highlighted).toBe(1);
  });
});

/**
 * highlightCurrentDay highlights the current day's task + score header
 * cells (row 4) with TODAY_HIGHLIGHT, after clearing any prior fill in
 * the header row.
 */
describe('highlightCurrentDay', () => {
  beforeEach(() => { state.weekly.currentDayIndex = 2; }); // Wednesday

  it('highlights G4 (task) and H4 (score) for Wednesday', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentDay(ctx, sheet);
    });

    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'G4')).toBe(CONFIG.COLORS.TODAY_HIGHLIGHT);
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'H4')).toBe(CONFIG.COLORS.TODAY_HIGHLIGHT);
  });

  it('clears prior header-row fill before applying new highlight', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    // Seed a stale highlight on Monday's columns.
    fake.helpers.setFill(CONFIG.WEEKLY_SHEET, 'C4', CONFIG.COLORS.TODAY_HIGHLIGHT);
    fake.helpers.setFill(CONFIG.WEEKLY_SHEET, 'D4', CONFIG.COLORS.TODAY_HIGHLIGHT);
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await highlightCurrentDay(ctx, sheet);
    });

    // Old Monday highlights cleared.
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'C4')).toBeNull();
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'D4')).toBeNull();
    // New Wednesday highlights present.
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'G4')).toBe(CONFIG.COLORS.TODAY_HIGHLIGHT);
  });
});
