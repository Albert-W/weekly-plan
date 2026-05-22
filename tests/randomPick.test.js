import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state } from './harness.js';

/**
 * randomPick tests.
 *
 * Behavior:
 *  - Fills empty task cells for the current day, only where the
 *    corresponding time-column cell has a value.
 *  - Leaves rows that already have a task selected untouched.
 *  - No-op if Tasks sheet is missing or empty.
 */

const randomPick = globalThis.randomPick;

function setupRandomPickSheets({ tasks = [], existingTasksForCurrentDay = {} } = {}) {
  const fake = makeFakeExcel({
    sheets: [CONFIG.WEEKLY_SHEET, CONFIG.TASKS_SHEET],
    activeSheet: CONFIG.WEEKLY_SHEET,
  });

  // Tasks sheet: headers in A1-A3, tasks from row 4 (mimics real layout).
  fake.helpers.setCells(CONFIG.TASKS_SHEET, { A1: 'h', A2: 'h', A3: 'h' });
  tasks.forEach((t, i) => {
    fake.helpers.setCells(CONFIG.TASKS_SHEET, { [`A${4 + i}`]: t });
  });
  state.weekly.lastTaskRow = tasks.length ? 3 + tasks.length : 3;

  // Time labels in column B for rows 5..10
  for (let row = 5; row <= 10; row++) {
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { [`B${row}`]: 8 + (row - 5) });
  }
  // Add any existing task entries
  for (const [addr, value] of Object.entries(existingTasksForCurrentDay)) {
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { [addr]: value });
  }

  state.weekly.currentDayIndex = 0; // Monday => task column C
  state.weekly.lastTimeRow = 10;

  return fake;
}

describe('randomPick', () => {
  beforeEach(() => {
    state.weekly.currentDayIndex = 0;
    state.weekly.lastTimeRow = 36;
  });

  it('fills empty task cells where the time column has a value', async () => {
    const fake = setupRandomPickSheets({ tasks: ['Read', 'Walk', 'Code'] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await randomPick(ctx); });

    // All 6 rows (5..10) should now have a task in column C
    for (let row = 5; row <= 10; row++) {
      const v = fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `C${row}`);
      expect(['Read', 'Walk', 'Code']).toContain(v);
    }
  });

  it('leaves pre-existing tasks untouched', async () => {
    const fake = setupRandomPickSheets({
      tasks: ['Read', 'Walk'],
      existingTasksForCurrentDay: { C7: 'KeepMe' },
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await randomPick(ctx); });

    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C7')).toBe('KeepMe');
    // Surrounding rows still got filled
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C5')).not.toBeNull();
  });

  it('does nothing for rows with no time-column value', async () => {
    const fake = setupRandomPickSheets({ tasks: ['Read'] });
    // Wipe the time value in row 8 to make it empty
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B8: '' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await randomPick(ctx); });

    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C8')).toBeNull();
  });

  it('handles current day = Tuesday by writing to column E', async () => {
    const fake = setupRandomPickSheets({ tasks: ['T'] });
    state.weekly.currentDayIndex = 1; // Tuesday => task column E
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await randomPick(ctx); });

    // Column C (Monday) untouched; column E (Tuesday) filled
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C5')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'E5')).toBe('T');
  });

  it('no-ops when the Tasks sheet has zero tasks', async () => {
    const fake = setupRandomPickSheets({ tasks: [] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await randomPick(ctx); });

    for (let row = 5; row <= 10; row++) {
      expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `C${row}`)).toBeNull();
    }
  });
});
