import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state } from './harness.js';

/**
 * Background time-highlight ticker tests.
 *
 * The ticker calls tickTimeHighlight() every minute. We don't test
 * the timer scheduling itself (would need fake timers + Excel and
 * has too many moving parts); we test the routing logic — that the
 * tick is a no-op when the user isn't currently on the Weekly sheet.
 */

describe('tickTimeHighlight', () => {
  beforeEach(() => {
    state.weekly.currentDayIndex = 0;
    state.weekly.lastMonday = new Date(2024, 0, 1);
  });

  it('does NOT touch Excel when currentSheet is not Weekly', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET, CONFIG.HABITS_SHEET],
      activeSheet: CONFIG.HABITS_SHEET,
    });
    fake.installAsExcelGlobal();
    state.currentSheet = CONFIG.HABITS_SHEET;
    fake.helpers.resetSyncCount();

    await globalThis.tickTimeHighlight();

    // Zero syncs because we never entered Excel.run.
    expect(fake.helpers.getSyncCount()).toBe(0);
  });

  it('runs Excel work when currentSheet IS Weekly', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET, CONFIG.HABITS_SHEET],
      activeSheet: CONFIG.WEEKLY_SHEET,
    });
    fake.installAsExcelGlobal();
    state.currentSheet = CONFIG.WEEKLY_SHEET;
    // Seed the time column so highlightCurrentTimeRow has something to look at.
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B5: 8 / 24 });
    fake.helpers.resetSyncCount();

    await globalThis.tickTimeHighlight();

    expect(fake.helpers.getSyncCount()).toBeGreaterThan(0);
  });
});
