import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, handleWeeklySelection } from './harness.js';

const handleHabitsSelection = globalThis.handleHabitsSelection;

/**
 * Selection-routing tests for the two domain modules' onSelection
 * callbacks. The registry plumbing is already covered by
 * registry.test.js; here we verify what the callbacks DO with the
 * dispatched address.
 */

describe('handleHabitsSelection', () => {
  beforeEach(() => {
    state.habits.lastRow = 10;
    state.habits.currentDayIndex = 0;
  });

  it('records a habit when column A is clicked in the data area', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET, CONFIG.SUMMARY_SHEET] });
    const row = CONFIG.HABITS.DATA_START_ROW;
    fake.helpers.setCells(CONFIG.HABITS_SHEET, {
      [`A${row}`]: 'Read',
      [`C${row}`]: 1,
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await handleHabitsSelection(ctx, `A${row}`, 'A', 1, row);
    });

    // recordHabitDone wrote +1 to D{row} (today)
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBe(1);
  });

  it('does NOT record when column A is clicked above the data area', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET, CONFIG.SUMMARY_SHEET] });
    // Click A1 — above DATA_START_ROW
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await handleHabitsSelection(ctx, 'A1', 'A', 1, 1);
    });

    // No write should have happened to any habit row.
    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, 'D4')).toBeNull();
  });

  it('ignores clicks in other columns', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.HABITS_SHEET, CONFIG.SUMMARY_SHEET] });
    const row = CONFIG.HABITS.DATA_START_ROW;
    fake.helpers.setCells(CONFIG.HABITS_SHEET, { [`A${row}`]: 'Read' });
    fake.installAsExcelGlobal();

    // Click on C{row} (base score column) — no action expected
    await Excel.run(async (ctx) => {
      await handleHabitsSelection(ctx, `C${row}`, 'C', 3, row);
    });

    expect(fake.helpers.getCellValue(CONFIG.HABITS_SHEET, `D${row}`)).toBeNull();
  });
});

describe('handleWeeklySelection', () => {
  beforeEach(() => {
    state.weekly.lastTimeRow = 36;
    state.weekly.scoreRow = 38;
  });

  it('shows an info status when a task column is clicked in the data area', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET, CONFIG.TASKS_SHEET] });
    fake.installAsExcelGlobal();

    let el = document.getElementById('status') || document.createElement('div');
    el.id = 'status'; el.textContent = ''; document.body.appendChild(el);

    await Excel.run(async (ctx) => {
      // Click C5 (Mon task, row 5)
      await handleWeeklySelection(ctx, 'C5', 'C', 3, 5);
    });

    expect(document.getElementById('status').textContent).toMatch(/Select a task/);
  });

  it('warns when a score column is clicked without a task selected', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET, CONFIG.TASKS_SHEET] });
    fake.installAsExcelGlobal();

    let el = document.getElementById('status') || document.createElement('div');
    el.id = 'status'; el.textContent = ''; document.body.appendChild(el);

    await Excel.run(async (ctx) => {
      // Click D5 (Mon score) — adjacent C5 has no task
      await handleWeeklySelection(ctx, 'D5', 'D', 4, 5);
    });

    // The modal shows a warning via showWarningPopup which calls showStatus.
    expect(document.getElementById('status').textContent).toMatch(/Please select a task/);
  });

  it('warns "already has score" when the score cell is already populated', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET, CONFIG.TASKS_SHEET] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Read', D5: 0.8 });
    fake.installAsExcelGlobal();

    let el = document.getElementById('status') || document.createElement('div');
    el.id = 'status'; el.textContent = ''; document.body.appendChild(el);

    await Excel.run(async (ctx) => {
      await handleWeeklySelection(ctx, 'D5', 'D', 4, 5);
    });

    expect(document.getElementById('status').textContent).toMatch(/can't be modified/i);
  });
});
