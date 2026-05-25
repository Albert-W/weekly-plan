import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, state, formatDateYYYYMMDD, getTodayScore, updateSummary } from './harness.js';

const refreshTodayScoreWidget = globalThis.refreshTodayScoreWidget;

// Pin clock — both setup and getTodayScore call formatDateYYYYMMDD
// (new Date()); a midnight crossing between the two would produce a
// today-row-not-found false negative. Task #38.
const FAKE_NOW = new Date(2024, 0, 1, 15, 30);
beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(FAKE_NOW); });
afterEach(() => { vi.useRealTimers(); });

/**
 * Today's-Score widget tests (task #26).
 *
 * Two layers:
 *   - getTodayScore() in summary.js: pure reader returning
 *     { positive, negative, total } | null.
 *   - refreshTodayScoreWidget() in weekly.js: orchestrator that
 *     drops getTodayScore output into three DOM cells.
 */

function setupSummary({ withTodayRow = false, positive = 0, negative = 0 } = {}) {
  const fake = makeFakeExcel({ sheets: [CONFIG.SUMMARY_SHEET] });
  state.weekly.lastSummaryRow = 0;
  fake.installAsExcelGlobal();
  if (withTodayRow) {
    const today = formatDateYYYYMMDD(new Date());
    fake.helpers.setCells(CONFIG.SUMMARY_SHEET, {
      A1: today,
      D1: positive,
      E1: negative,
      F1: positive + negative,
    });
    state.weekly.lastSummaryRow = 1;
  }
  return fake;
}

function setupScoreWidgetDom() {
  document.body.innerHTML = `
    <div id="today-score">
      <div id="ts-pos">—</div>
      <div id="ts-total">—</div>
      <div id="ts-neg">—</div>
    </div>
  `;
}

describe('getTodayScore', () => {
  beforeEach(() => { state.weekly.lastSummaryRow = 0; });

  it('returns null when today is not yet recorded', async () => {
    setupSummary();
    let result;
    await Excel.run(async (ctx) => { result = await getTodayScore(ctx); });
    expect(result).toBeNull();
  });

  it('returns the three numbers from today row when present', async () => {
    setupSummary({ withTodayRow: true, positive: 1.5, negative: -0.3 });
    let result;
    await Excel.run(async (ctx) => { result = await getTodayScore(ctx); });
    expect(result.positive).toBeCloseTo(1.5, 6);
    expect(result.negative).toBeCloseTo(-0.3, 6);
    expect(result.total).toBeCloseTo(1.2, 6);
  });

  it('returns null when the Summary sheet does not exist', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.installAsExcelGlobal();
    let result;
    await Excel.run(async (ctx) => { result = await getTodayScore(ctx); });
    expect(result).toBeNull();
  });

  it('integrates with updateSummary (round-trip)', async () => {
    const fake = setupSummary();
    await Excel.run(async (ctx) => { await updateSummary(ctx, 0.8, -0.2); });
    let result;
    await Excel.run(async (ctx) => { result = await getTodayScore(ctx); });
    expect(result.positive).toBeCloseTo(0.8, 6);
    expect(result.negative).toBeCloseTo(-0.2, 6);
    expect(result.total).toBeCloseTo(0.6, 6);
  });
});

describe('refreshTodayScoreWidget', () => {
  beforeEach(() => {
    setupScoreWidgetDom();
    state.weekly.lastSummaryRow = 0;
  });

  it('renders em-dashes when there is no recorded score for today', async () => {
    setupSummary();
    await refreshTodayScoreWidget();
    expect(document.getElementById('ts-pos').textContent).toBe('—');
    expect(document.getElementById('ts-total').textContent).toBe('—');
    expect(document.getElementById('ts-neg').textContent).toBe('—');
  });

  it('renders the three formatted score values when today exists', async () => {
    setupSummary({ withTodayRow: true, positive: 1.5, negative: -0.3 });
    await refreshTodayScoreWidget();
    expect(document.getElementById('ts-pos').textContent).toBe('1.5');
    // total = 1.2
    expect(document.getElementById('ts-total').textContent).toBe('1.2');
    expect(document.getElementById('ts-neg').textContent).toBe('-0.3');
  });

  it('is a no-op when the widget elements are missing from the DOM', async () => {
    document.body.innerHTML = ''; // no widget
    setupSummary({ withTodayRow: true, positive: 1, negative: 0 });
    // Should not throw.
    await expect(refreshTodayScoreWidget()).resolves.toBeUndefined();
  });

  it('single-flight: overlapping calls share one Excel.run (regression guard for task #36)', async () => {
    const fake = setupSummary({ withTodayRow: true, positive: 1.5, negative: -0.3 });
    fake.helpers.resetSyncCount();

    // Fire 5 calls without awaiting between them — simulates the
    // 60s ticker firing while the previous tick + score-change
    // refresh are still settling. Without the single-flight guard
    // each call would spawn its own Excel.run.
    const promises = [];
    for (let i = 0; i < 5; i++) promises.push(refreshTodayScoreWidget());
    await Promise.all(promises);

    // getTodayScore is 3 syncs (existence + date-col load + cell
    // loads); 5 unguarded calls would be 15. With the guard, only
    // the first call's run executes. (getTodayScore could itself be
    // batched to 2 syncs in a follow-up, mirroring task #37.)
    expect(fake.helpers.getSyncCount()).toBe(3);
    // Final DOM state still correct.
    expect(document.getElementById('ts-pos').textContent).toBe('1.5');
  });
});
