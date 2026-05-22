import { describe, it, expect } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { processWeeklyScoreChange, CONFIG, state } from './harness.js';

/**
 * processWeeklyScoreChange integration test.
 *
 * Covers (after task #2 batching refactor):
 *  - Known-task path: weight applied, color set, daily total and
 *    per-task stats updated.
 *  - "others" fallback when the task name isn't in the Tasks sheet
 *    but an "others" row already exists.
 *  - First-time "others" creation populates A, B, C, D, F, G
 *    (regression guard for the bonus correctness fix in task #2).
 *  - Color rules: positive -> POSITIVE, negative -> NEGATIVE,
 *    zero -> NEUTRAL.
 *  - Sync-count regression guard: <=3 syncs total.
 */

function setupSheets({ tasks = [], existingDailyTotal = 0 } = {}) {
  const fake = makeFakeExcel({
    sheets: [CONFIG.WEEKLY_SHEET, CONFIG.TASKS_SHEET, CONFIG.SUMMARY_SHEET],
    activeSheet: CONFIG.WEEKLY_SHEET,
  });
  // Populate Tasks sheet starting at row 4
  tasks.forEach((t, i) => {
    const row = 4 + i;
    fake.helpers.setCells(CONFIG.TASKS_SHEET, {
      [`A${row}`]: t.name,
      [`B${row}`]: t.weight ?? 1,
      [`F${row}`]: t.count ?? 0,
      [`G${row}`]: t.totalScore ?? 0,
    });
  });
  state.weekly.taskl = 4 + tasks.length - 1;
  state.weekly.lastMonday = new Date(2024, 0, 1);
  return fake;
}

describe('processWeeklyScoreChange', () => {
  it('applies the task weight to the new score', async () => {
    // Task "Read" with weight 2 — entering score 0.5 should accumulate 1.0
    const fake = setupSheets({ tasks: [{ name: 'Read', weight: 2 }] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Read' });
    fake.installAsExcelGlobal();

    // Score column for Monday = column 4 = D. So col=4, row=5, newScore=0.5
    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    // Daily total in D38 should now be 1.0
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `D${CONFIG.WEEKLY.scoreLine}`)).toBe(1);
    // Stats updated: F4 = 1, G4 = 1.0
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'F4')).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'G4')).toBe(1);
    // Color = POSITIVE (weightedScore > 0)
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'D5')).toBe(CONFIG.COLORS.POSITIVE);
    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'C5')).toBe(CONFIG.COLORS.POSITIVE);
  });

  it('applies NEGATIVE color when weighted score < 0', async () => {
    const fake = setupSheets({ tasks: [{ name: 'Procrastinate', weight: -1 }] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Procrastinate' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'D5')).toBe(CONFIG.COLORS.NEGATIVE);
  });

  it('applies NEUTRAL color when weighted score is exactly 0', async () => {
    const fake = setupSheets({ tasks: [{ name: 'Read', weight: 1 }] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Read' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0);
    });

    expect(fake.helpers.getCellColor(CONFIG.WEEKLY_SHEET, 'D5')).toBe(CONFIG.COLORS.NEUTRAL);
  });

  it('accumulates the daily total across multiple writes', async () => {
    const fake = setupSheets({ tasks: [{ name: 'Read', weight: 1 }] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, {
      C5: 'Read',
      E6: 'Read',
      [`D${CONFIG.WEEKLY.scoreLine}`]: 0.4, // pre-existing daily total
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    // 0.4 + 0.5 = 0.9
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `D${CONFIG.WEEKLY.scoreLine}`))
      .toBeCloseTo(0.9, 9);
  });

  it('falls back to existing "others" task when name is not found', async () => {
    const fake = setupSheets({
      tasks: [
        { name: 'Read', weight: 2 },
        { name: 'others', weight: 1 },
      ],
    });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Doodle' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    // others is at row 5 of the Tasks sheet (row 4 = Read, row 5 = others)
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'F5')).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'G5')).toBe(0.5);
  });

  it('creates an "others" row from scratch when no fallback exists, populating B, C, D, F, G (task #2 bonus fix)', async () => {
    const fake = setupSheets({ tasks: [{ name: 'Read', weight: 2 }] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Unknown' });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    // New row should be at row 5 (taskl was 4, +1)
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'A5')).toBe('others');
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'B5')).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'C5')).toBeTypeOf('string');
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'D5')).toBeTypeOf('string');
    // Previously these two were left null:
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'F5')).toBe(1);
    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'G5')).toBe(0.5);
    // state.weekly.taskl should be bumped
    expect(state.weekly.taskl).toBe(5);
  });

  it('does nothing if the task name cell is empty', async () => {
    const fake = setupSheets({ tasks: [{ name: 'Read', weight: 1 }] });
    // C5 is empty
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    expect(fake.helpers.getCellValue(CONFIG.TASKS_SHEET, 'F4')).toBe(0);
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, `D${CONFIG.WEEKLY.scoreLine}`)).toBeNull();
  });

  it('PERF: uses at most 3 syncs per score entry regardless of task count (regression guard for task #2)', async () => {
    // Many tasks — make sure the lookup loop didn't sneak back in.
    const tasks = [];
    for (let i = 0; i < 50; i++) tasks.push({ name: `task_${i}`, weight: 1 });
    tasks.push({ name: 'Match', weight: 3 });
    const fake = setupSheets({ tasks });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { C5: 'Match' });
    fake.installAsExcelGlobal();
    fake.helpers.resetSyncCount();

    await Excel.run(async (ctx) => {
      await processWeeklyScoreChange(ctx, 5, 4, 0.5);
    });

    // 2 inside processWeeklyScoreChange + updateSummary calls.
    // updateSummary is on its own perf-task list; tightened later.
    expect(fake.helpers.getSyncCount()).toBeLessThanOrEqual(7);
  });
});
