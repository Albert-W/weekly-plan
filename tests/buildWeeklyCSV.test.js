import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { buildWeeklyCSV, CONFIG } from './harness.js';

/**
 * Integration tests for buildWeeklyCSV.
 * Verifies the post-refactor CSV format matches what the production
 * code produced before task #14 split it into a helper.
 */

function setupWeeklySheet() {
  const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET], activeSheet: CONFIG.WEEKLY_SHEET });
  const { helpers } = fake;
  // B4 = "yyyy mm" label
  helpers.setCells(CONFIG.WEEKLY_SHEET, {
    B4: '2024 01',
    // Day headers D4:P4 = 1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 6, 0, 7
    // (task columns hold the day numbers, score columns are blank)
    D4: 1, E4: '', F4: 2, G4: '', H4: 3, I4: '', J4: 4,
    K4: '', L4: 5, M4: '', N4: 6, O4: '', P4: 7,
  });
  // Time column (B5..)
  helpers.setCells(CONFIG.WEEKLY_SHEET, {
    B5: 8 / 24,        // 08:00 as fraction-of-day
    B6: 9 / 24,        // 09:00
    B7: '',            // empty — should be skipped
    B8: 10 / 24,       // 10:00
  });
  // A scored Monday entry at 08:00 + an unscored Wednesday entry at 09:00
  helpers.setCells(CONFIG.WEEKLY_SHEET, {
    C5: 'Read',   // Mon task
    D5: 0.8,      // Mon score
    G6: 'Walk',   // Wed task
    H6: '',       // Wed score (empty)
  });
  return fake;
}

describe('buildWeeklyCSV', () => {
  it('returns { csv, filename } with the expected filename format', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    expect(result.filename).toBe('Weekly_2024-01_1-7.csv');
  });

  it('emits a "Time" + 14-column header row', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    const lines = result.csv.split('\n');
    expect(lines[0]).toBe(
      'Time,Mon_Task,Mon_Score,Tue_Task,Tue_Score,Wed_Task,Wed_Score,Thu_Task,Thu_Score,Fri_Task,Fri_Score,Sat_Task,Sat_Score,Sun_Task,Sun_Score'
    );
  });

  it('formats fractional time values as HH:MM', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    expect(result.csv).toContain('08:00,');
    expect(result.csv).toContain('09:00,');
    expect(result.csv).toContain('10:00,');
  });

  it('skips rows where the time cell is empty', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    const lines = result.csv.split('\n').filter(Boolean);
    // header + 3 non-empty time rows = 4 lines
    expect(lines.length).toBe(4);
  });

  it('emits task and score values in the correct columns', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    const lines = result.csv.split('\n');
    // Mon col = index 1 (task) and 2 (score)
    const monRow = lines.find((l) => l.startsWith('08:00'));
    expect(monRow).toContain(',Read,0.8,');

    // Wed task with empty score
    const wedRow = lines.find((l) => l.startsWith('09:00'));
    expect(wedRow).toContain(',Walk,,');
  });

  it('returns CSV ending with a trailing newline (matches legacy output)', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();

    let result;
    await Excel.run(async (ctx) => {
      result = await buildWeeklyCSV(ctx);
    });

    expect(result.csv.endsWith('\n')).toBe(true);
  });

  it('uses a single context.sync() round-trip for all reads (regression guard for the perf hot path)', async () => {
    const fake = setupWeeklySheet();
    fake.installAsExcelGlobal();
    fake.helpers.resetSyncCount();

    await Excel.run(async (ctx) => {
      await buildWeeklyCSV(ctx);
    });

    expect(fake.helpers.getSyncCount()).toBe(1);
  });
});
