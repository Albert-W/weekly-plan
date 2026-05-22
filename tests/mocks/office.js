/**
 * Minimal Office.js stub installed as a global so production code can
 * call `Office.onReady(...)`, check `Office.HostType.Excel`, etc.
 * without throwing in a Node/jsdom test environment.
 *
 * Tests don't trigger Office.onReady; we only need the symbols to
 * exist so the src files can load without errors.
 */

export function installOfficeGlobal(globalObj = globalThis) {
  globalObj.Office = {
    onReady: () => {},
    HostType: { Excel: 'Excel' },
    PlatformType: { OfficeOnline: 'OfficeOnline', PC: 'PC', Mac: 'Mac' },
    context: {
      platform: 'PC',
      ui: {},
    },
  };
}
