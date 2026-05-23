import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, addTask } from './harness.js';

/**
 * addTask is the UI form wrapper. Reads the form inputs, calls
 * createTask (already tested in createTask.test.js), then clears
 * the form and shows a status banner. Tests:
 *  - Happy path: full form -> task created, form cleared, success
 *  - Empty name -> warning, no Excel work attempted
 *  - Bad weight -> defaults to 1
 */

function setupDom() {
  document.body.innerHTML = `
    <input id="new-task-name" type="text" />
    <input id="new-task-weight" type="number" value="1" />
    <div id="status"></div>
  `;
}

function setupTasksFake() {
  const fake = makeFakeExcel({ sheets: [CONFIG.TASKS_SHEET] });
  // Real Tasks sheet has header rows 1-3 (createTask uses
  // usedRange.rowCount as the last-row number).
  fake.helpers.setCells(CONFIG.TASKS_SHEET, { A1: 'h', A2: 'h', A3: 'h' });
  state.weekly.lastTaskRow = 3;
  return fake;
}

describe('addTask (UI form wrapper)', () => {
  beforeEach(() => { setupDom(); });

  it('happy path: creates the task, clears the form, shows success', async () => {
    const fake = setupTasksFake();
    fake.installAsExcelGlobal();
    document.getElementById('new-task-name').value = 'Deep Work';
    document.getElementById('new-task-weight').value = '2';

    await addTask();

    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'A4')).toBe('Deep Work');
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'B4')).toBe(2);
    expect(document.getElementById('new-task-name').value).toBe('');
    expect(document.getElementById('new-task-weight').value).toBe('1');
    expect(document.getElementById('status').textContent).toMatch(/added/);
  });

  it('refuses to create with empty name and warns the user', async () => {
    const fake = setupTasksFake();
    fake.installAsExcelGlobal();
    document.getElementById('new-task-name').value = '   ';

    await addTask();

    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'A4')).toBeNull();
    expect(document.getElementById('status').textContent).toMatch(/enter a task name/i);
  });

  it('falls back to weight=1 when the input is non-numeric', async () => {
    const fake = setupTasksFake();
    fake.installAsExcelGlobal();
    document.getElementById('new-task-name').value = 'foo';
    document.getElementById('new-task-weight').value = 'abc';

    await addTask();

    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'B4')).toBe(1);
  });
});
