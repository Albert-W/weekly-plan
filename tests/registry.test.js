import { describe, it, expect, beforeEach } from 'vitest';
import { CONFIG } from './harness.js';

/**
 * Registry tests for task #19 — sheet-handler routing.
 * Confirms:
 *   - Domain modules (habits, weekly) self-registered on load.
 *   - Adding a new sheet handler requires no edits to events.js.
 */

describe('sheet-handler registry', () => {
  it('habits.js registered an onSelection handler for the Habits sheet', () => {
    const h = globalThis.getSheetHandlers(CONFIG.HABITS_SHEET);
    expect(h).not.toBeNull();
    expect(typeof h.onSelection).toBe('function');
  });

  it('weekly.js registered onSelection, onActivate, and onChange for the Weekly sheet', () => {
    const h = globalThis.getSheetHandlers(CONFIG.WEEKLY_SHEET);
    expect(h).not.toBeNull();
    expect(typeof h.onSelection).toBe('function');
    expect(typeof h.onActivate).toBe('function');
    expect(typeof h.onChange).toBe('function');
  });

  it('returns null for an unregistered sheet name', () => {
    expect(globalThis.getSheetHandlers('NoSuchSheet')).toBeNull();
  });

  it('registerSheetHandlers is composable: a second call merges into the first', () => {
    const sheet = '__test_compose__';
    globalThis.registerSheetHandlers(sheet, { onActivate: () => 'A' });
    globalThis.registerSheetHandlers(sheet, { onSelection: () => 'B' });
    const h = globalThis.getSheetHandlers(sheet);
    expect(h.onActivate()).toBe('A');
    expect(h.onSelection()).toBe('B');
  });
});
