import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { createTask, CONFIG, state } from './harness.js';

/**
 * createTask domain function tests (task #16).
 * Verifies the Excel-side behavior in isolation from the DOM form.
 */

function setupTasksSheet({ existingTasks = [] } = {}) {
  const fake = makeFakeExcel({ sheets: [CONFIG.TASKS_SHEET], activeSheet: CONFIG.TASKS_SHEET });
  // Mimic the real Tasks sheet: header rows in A1-A3 so that
  // getUsedRange().rowCount == last populated row number (which is
  // what the production code assumes).
  fake.helpers.setCells(CONFIG.TASKS_SHEET, {
    A1: 'Header',
    A2: 'Header',
    A3: 'Header',
  });
  existingTasks.forEach((t, i) => {
    const row = 4 + i;
    fake.helpers.setCells(CONFIG.TASKS_SHEET, {
      [`A${row}`]: t.name,
      [`B${row}`]: t.weight ?? 1,
    });
  });
  state.weekly.lastTaskRow = existingTasks.length ? 4 + existingTasks.length - 1 : 3;
  return fake;
}

describe('createTask', () => {
  beforeEach(() => {
    state.weekly.lastTaskRow = 3;
  });

  it('appends a row with name, weight, and timestamp', async () => {
    const fake = setupTasksSheet({ existingTasks: [{ name: 'Read', weight: 1 }] });
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await createTask(ctx, 'Deep Work', 2);
    });

    // Existing row was at row 4 -> new row at row 5
    expect(result.row).toBe(5);
    expect(result.name).toBe('Deep Work');
    expect(result.weight).toBe(2);

    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'A5')).toBe('Deep Work');
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'B5')).toBe(2);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'C5')).toBeTypeOf('string');
  });

  it('updates state.weekly.lastTaskRow to the new row', async () => {
    const fake = setupTasksSheet({ existingTasks: [{ name: 'Read' }, { name: 'Walk' }] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await createTask(ctx, 'New', 1); });

    expect(state.weekly.lastTaskRow).toBe(6);
  });

  it('trims whitespace from the name', async () => {
    const fake = setupTasksSheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await createTask(ctx, '  Padded  ', 1);
    });

    expect(result.name).toBe('Padded');
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, `A${result.row}`)).toBe('Padded');
  });

  it('defaults weight to 1 when given a non-finite value', async () => {
    const fake = setupTasksSheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await createTask(ctx, 'foo', NaN);
    });

    expect(result.weight).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, `B${result.row}`)).toBe(1);
  });

  it('rejects empty names', async () => {
    const fake = setupTasksSheet();
    fake.installAsExcelGlobal();

    await expect(
      Excel.run(async (ctx) => { await createTask(ctx, '   ', 1); })
    ).rejects.toThrow(/required/i);
  });

  it('throws when the Tasks sheet does not exist', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] }); // no Tasks
    fake.installAsExcelGlobal();

    await expect(
      Excel.run(async (ctx) => { await createTask(ctx, 'foo', 1); })
    ).rejects.toThrow(/not found/i);
  });
});
