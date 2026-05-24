import { describe, it, expect, beforeEach } from 'vitest';
import { makeFakeExcel } from './mocks/excel.js';
import {
  registerOnChangedEvent, CONFIG, state,
  registerSheetHandlers, resetWriteChains,
} from './harness.js';

const handleCellChanged = globalThis.handleCellChanged;

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

/**
 * Race regression guard for task #32.
 *
 * Before the fix: two rapid handleCellChanged invocations would
 * interleave through the queued-handler awaits; the load-then-write
 * pattern inside the dispatched handler would let increments
 * clobber each other.
 *
 * Strategy: register a fake onChange handler that DELIBERATELY does
 * a read-modify-write in the racy shape. Fire two events without
 * awaiting between them. Assert the final value reflects BOTH
 * increments.
 */
describe('handleCellChanged serialization (task #32)', () => {
  beforeEach(() => {
    resetWriteChains();
    state.currentSheet = CONFIG.WEEKLY_SHEET;
  });

  it('serializes two overlapping handleCellChanged calls for the same sheet', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET],
      activeSheet: CONFIG.WEEKLY_SHEET,
    });
    // Seed the running-total cell.
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { Z1: 0 });
    fake.installAsExcelGlobal();

    // RMW handler: load Z1, add 1, write back. Without serialization
    // two concurrent calls will both read 0, both write 1.
    const rmwHandler = async (context /*, address, colIndex, row */) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      const cell = sheet.getRange('Z1');
      cell.load('values');
      await context.sync();
      const cur = parseFloat(cell.values[0][0]) || 0;
      // Force a yield BEFORE the write so a concurrent run can race
      // here if serialization were missing.
      await Promise.resolve();
      cell.values = [[cur + 1]];
      await context.sync();
    };

    registerSheetHandlers(CONFIG.WEEKLY_SHEET, { onChange: rmwHandler });

    // Fire two events without awaiting between them.
    const pA = handleCellChanged({ address: 'D5' });
    const pB = handleCellChanged({ address: 'D6' });
    await Promise.all([pA, pB]);

    // Both increments preserved -> value should be 2.
    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'Z1')).toBe(2);
  });

  it('preserves all five increments when five events fire concurrently', async () => {
    const fake = makeFakeExcel({
      sheets: [CONFIG.WEEKLY_SHEET],
      activeSheet: CONFIG.WEEKLY_SHEET,
    });
    fake.helpers.setCells(CONFIG.WEEKLY_SHEET, { Z2: 0 });
    fake.installAsExcelGlobal();

    const rmw = async (context) => {
      const sheet = context.workbook.worksheets.getItem(CONFIG.WEEKLY_SHEET);
      const cell = sheet.getRange('Z2');
      cell.load('values');
      await context.sync();
      const cur = parseFloat(cell.values[0][0]) || 0;
      await Promise.resolve();
      cell.values = [[cur + 1]];
      await context.sync();
    };
    registerSheetHandlers(CONFIG.WEEKLY_SHEET, { onChange: rmw });

    const promises = [];
    for (let i = 0; i < 5; i++) {
      promises.push(handleCellChanged({ address: `D${5 + i}` }));
    }
    await Promise.all(promises);

    expect(fake.helpers.getCellValue(CONFIG.WEEKLY_SHEET, 'Z2')).toBe(5);
  });
});
