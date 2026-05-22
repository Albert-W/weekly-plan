import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import { registerOnChangedEvent, CONFIG, state } from './harness.js';

/**
 * Regression guard for task #13.
 *
 * Before the fix, registerOnChangedEvent stacked new handlers on
 * every call without removing the previous one. After the fix, it
 * tracks the handler in state.changeHandler and removes it before
 * adding a new one.
 *
 * Strategy: register twice on the same sheet; assert that only ONE
 * handler is currently attached on the fake sheet.
 */

describe('registerOnChangedEvent', () => {
  beforeEach(() => {
    // Make sure state from previous test doesn't leak in.
    state.changeHandler = null;
  });

  it('removes the previous handler before adding a new one', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET],
      activeSheet: CONFIG.WEEKLY_SHEET,
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await registerOnChangedEvent(ctx, sheet);
    });
    expect(fake.helpers.getHandlerCount(CONFIG.WEEKLY_SHEET, 'onChanged')).toBe(1);

    // Re-register. Without the fix, this would leave 2 handlers.
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await registerOnChangedEvent(ctx, sheet);
    });
    expect(fake.helpers.getHandlerCount(CONFIG.WEEKLY_SHEET, 'onChanged')).toBe(1);

    // And again — still one.
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await registerOnChangedEvent(ctx, sheet);
    });
    expect(fake.helpers.getHandlerCount(CONFIG.WEEKLY_SHEET, 'onChanged')).toBe(1);
  });

  it('stores the active handler on state.changeHandler', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET],
      activeSheet: CONFIG.WEEKLY_SHEET,
    });
    fake.installAsExcelGlobal();

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      await registerOnChangedEvent(ctx, sheet);
    });

    expect(state.changeHandler).not.toBeNull();
    expect(state.changeHandler.removed).toBe(false);
  });
});
