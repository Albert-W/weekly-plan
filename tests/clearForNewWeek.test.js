import { describe, it, expect } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { clearForNewWeek, CONFIG } from './harness.js';

/**
 * clearForNewWeek integration test.
 *
 * Legacy behavior contract (preserved through task #15 batching):
 *  - Background fill on C5:Z{scoreLine} is reset.
 *  - The score totals row (row 38) is cleared.
 *  - Task/score pairs are cleared ONLY for rows where a score was set.
 *  - Rows with a task selected but no score are LEFT INTACT.
 *
 * Plus a perf regression guard: must complete in exactly 2 syncs.
 */

function setupSheet({ entries }) {
  const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET], activeSheet: CONFIG.WEEKLY_SHEET });
  for (const e of entries) {
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, {
      [`${e.taskCol}${e.row}`]: e.task,
    });
    if (e.score !== undefined) {
      fake.helpers.setCells(CONFIG.WEEKLY_SHEET, {
        [`${e.scoreCol}${e.row}`]: e.score,
      });
    }
    if (e.color) {
      fake.helpers.setFill(CONFIG.WEEKLY_SHEET, `${e.taskCol}${e.row}`, e.color);
      fake.helpers.setFill(CONFIG.WEEKLY_SHEET, `${e.scoreCol}${e.row}`, e.color);
    }
  }
  // Seed the score totals row with some numbers to confirm clearing
  fake.helpers.setCells(CONFIG.WEEKLY_SHEET, {
    [`C${CONFIG.WEEKLY.scoreLine}`]: 1.5,
    [`D${CONFIG.WEEKLY.scoreLine}`]: 2.0,
  });
  return fake;
}

describe('clearForNewWeek', () => {
  it('clears task+score for rows with a recorded score', async () => {
    const fake = setupSheet({
      entries: [
        { row: 5, taskCol: 'C', scoreCol: 'D', task: 'Read', score: 0.8, color: '#abc' },
        { row: 10, taskCol: 'E', scoreCol: 'F', task: 'Walk', score: 0.6, color: '#def' },
      ],
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await clearForNewWeek(ctx); });

    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C5')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'D5')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'E10')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'F10')).toBeNull();
  });

  it('preserves rows with a task selected but no score (legacy VBA contract)', async () => {
    const fake = setupSheet({
      entries: [
        // Task entered but never scored
        { row: 7, taskCol: 'C', scoreCol: 'D', task: 'Read' },
        // Task + score
        { row: 8, taskCol: 'C', scoreCol: 'D', task: 'Walk', score: 0.4 },
      ],
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await clearForNewWeek(ctx); });

    // Untouched task survives
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C7')).toBe('Read');
    // Scored row was cleared
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'C8')).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'D8')).toBeNull();
  });

  it('clears the daily scores totals row', async () => {
    const fake = setupSheet({ entries: [] });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await clearForNewWeek(ctx); });

    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `C${CONFIG.WEEKLY.scoreLine}`)).toBeNull();
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `D${CONFIG.WEEKLY.scoreLine}`)).toBeNull();
  });

  it('resets background fill on the data area', async () => {
    const fake = setupSheet({
      entries: [
        { row: 5, taskCol: 'C', scoreCol: 'D', task: 'Read', score: 0.8, color: '#70AD47' },
      ],
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => { await clearForNewWeek(ctx); });

    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'C5')).toBeNull();
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'D5')).toBeNull();
  });

  it('PERF: completes in exactly 2 syncs regardless of how many cells are filled (regression guard for task #15)', async () => {
    // Fill 50 rows worth of scores across multiple days
    const entries = [];
    const taskCols = ['C', 'E', 'G', 'I', 'K', 'M', 'O'];
    const scoreCols = ['D', 'F', 'H', 'J', 'L', 'N', 'P'];
    for (let row = 5; row <= 30; row++) {
      for (let d = 0; d < 7; d++) {
        entries.push({
          row,
          taskCol: taskCols[d],
          scoreCol: scoreCols[d],
          task: 't',
          score: 0.5,
        });
      }
    }
    const fake = setupSheet({ entries });
    fake.installAsExcelGlobal();
    fake.helpers.resetSyncCount();

    await Excel.run(async (ctx) => { await clearForNewWeek(ctx); });

    expect(fake.helpers.getSyncCount()).toBe(2);
  });
});
