import { describe, it, expect, beforeEach, afterEach } from 'vitest';

/**
 * Platform-guard tests for tasks #22 and #23.
 *  - isExcelOnline() never throws even when Office.PlatformType is
 *    undefined (the bug task #23 was filed for).
 *  - downloadCSV no-ops the <a download> path on Desktop and shows
 *    a clear status banner instead.
 */

const isExcelOnline = globalThis.isExcelOnline;
const downloadCSV = globalThis.downloadCSV;

describe('isExcelOnline guard', () => {
  let originalPlatformType;

  beforeEach(() => {
    originalPlatformType = globalThis.Office.PlatformType;
  });

  afterEach(() => {
    globalThis.Office.PlatformType = originalPlatformType;
    globalThis.Office.context.platform = 'PC';
  });

  it('returns true on Excel Online', () => {
    globalThis.Office.context.platform = 'OfficeOnline';
    expect(isExcelOnline()).toBe(true);
  });

  it('returns false on Desktop', () => {
    globalThis.Office.context.platform = 'PC';
    expect(isExcelOnline()).toBe(false);
  });

  it('does NOT throw when Office.PlatformType is undefined (older clients)', () => {
    delete globalThis.Office.PlatformType;
    expect(() => isExcelOnline()).not.toThrow();
    expect(isExcelOnline()).toBe(false);
  });
});

describe('downloadCSV platform gating', () => {
  let originalPlatform;
  let originalAppend;
  let appended = [];

  beforeEach(() => {
    originalPlatform = globalThis.Office.context.platform;
    appended = [];
    // jsdom doesn't ship URL.createObjectURL; stub it.
    if (typeof URL.createObjectURL !== 'function') {
      URL.createObjectURL = () => 'blob:fake';
      URL.revokeObjectURL = () => {};
    }
    // Spy on body.appendChild to detect the <a download> path.
    originalAppend = document.body.appendChild.bind(document.body);
    document.body.appendChild = (node) => {
      if (node.tagName === 'A') appended.push(node);
      return originalAppend(node);
    };
    // Fresh status element.
    let el = document.getElementById('status');
    if (!el) {
      el = document.createElement('div');
      el.id = 'status';
      originalAppend(el);
    }
    el.textContent = '';
  });

  afterEach(() => {
    globalThis.Office.context.platform = originalPlatform;
    document.body.appendChild = originalAppend;
  });

  it('on Desktop: returns false, does not append an <a> link, shows a warning banner', () => {
    globalThis.Office.context.platform = 'PC';
    const ok = downloadCSV('a,b,c\n1,2,3\n', 'Test.csv');
    expect(ok).toBe(false);
    expect(appended.length).toBe(0);
    const banner = document.getElementById('status');
    expect(banner.textContent).toContain('Desktop');
    expect(banner.textContent).toContain('Test.csv');
  });

  it('on Excel Online: returns true and appends an <a> link briefly', () => {
    globalThis.Office.context.platform = 'OfficeOnline';
    const ok = downloadCSV('a,b\n1,2\n', 'Online.csv');
    expect(ok).toBe(true);
    // The link is appended then removed, so we look at our captured list.
    expect(appended.length).toBe(1);
    expect(appended[0].getAttribute('download')).toBe('Online.csv');
  });
});
