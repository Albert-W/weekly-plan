import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { CONFIG, refreshTimeHighlight } from './harness.js';

/**
 * refreshTimeHighlight tests.
 *
 * Wrapper over highlightCurrentTimeRow that:
 *  - Owns its own Excel.run
 *  - Returns early if the Weekly sheet doesn't exist (no throw)
 *  - Accepts a { silent: true } option to suppress the success banner
 */

describe('refreshTimeHighlight', () => {
  beforeEach(() => {
    let el = document.getElementById('status');
    if (!el) {
      el = document.createElement('div');
      el.id = 'status';
      document.body.appendChild(el);
    }
    el.textContent = '';
    el.className = 'status';
  });

  it('shows a success status banner by default', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B5: 8 });
    fake.installAsExcelGlobal();

    await refreshTimeHighlight();

    const banner = document.getElementById('status').textContent;
    expect(banner).toMatch(/refreshed/i);
  });

  it('does NOT show a banner when silent: true', async () => {
    const fake = makeFakeExcel({ sheets: [CONFIG.WEEKLY_SHEET] });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { B5: 8 });
    fake.installAsExcelGlobal();

    await refreshTimeHighlight({ silent: true });

    expect(document.getElementById('status').textContent).toBe('');
  });

  it('is a no-op when the Weekly sheet does not exist', async () => {
    const fake = makeFakeExcel({ sheets: ['SomeOtherSheet'] });
    fake.installAsExcelGlobal();

    // Should not throw
    await expect(refreshTimeHighlight({ silent: true })).resolves.not.toThrow();
  });
});
