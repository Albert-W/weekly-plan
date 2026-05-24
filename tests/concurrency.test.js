import { describe, it, expect, beforeEach } from 'vitest';
import {
  serializeSheetWrite, resetWriteChains,
} from './harness.js';

/**
 * Unit tests for the per-sheet write-serialization helper introduced
 * for task #32. Three properties to pin:
 *
 *   1. Two tasks queued for the same sheet name run sequentially.
 *   2. A thrown task does not poison the chain — subsequent tasks
 *      for the same sheet still run.
 *   3. Tasks for different sheet names run independently (can
 *      overlap).
 */

beforeEach(() => { resetWriteChains(); });

describe('serializeSheetWrite', () => {
  it('runs two tasks for the same sheet sequentially (no overlap)', async () => {
    const events = [];
    const slow = (label) => async () => {
      events.push(`start ${label}`);
      await new Promise((r) => setTimeout(r, 5));
      events.push(`end ${label}`);
      return label;
    };

    const pA = serializeSheetWrite('Sheet1', slow('A'));
    const pB = serializeSheetWrite('Sheet1', slow('B'));

    const [a, b] = await Promise.all([pA, pB]);
    expect(a).toBe('A');
    expect(b).toBe('B');
    // Exact interleaving — B must NOT start until A ends.
    expect(events).toEqual(['start A', 'end A', 'start B', 'end B']);
  });

  it('does NOT poison the chain when a task throws', async () => {
    const events = [];
    const ok = (label) => async () => { events.push(label); return label; };
    const bad = async () => {
      events.push('bad-start');
      throw new Error('boom');
    };

    const pA = serializeSheetWrite('Sheet1', ok('A'));
    const pBad = serializeSheetWrite('Sheet1', bad);
    const pC = serializeSheetWrite('Sheet1', ok('C'));

    await expect(pA).resolves.toBe('A');
    // The bad task rejects — the chain swallows it for downstream
    // continuity, but the original caller still sees the rejection.
    await expect(pBad).rejects.toThrow('boom');
    await expect(pC).resolves.toBe('C');

    expect(events).toEqual(['A', 'bad-start', 'C']);
  });

  it('runs tasks for DIFFERENT sheet names in parallel', async () => {
    const events = [];
    const slow = (label) => async () => {
      events.push(`start ${label}`);
      await new Promise((r) => setTimeout(r, 10));
      events.push(`end ${label}`);
    };

    const pA = serializeSheetWrite('SheetA', slow('A'));
    const pB = serializeSheetWrite('SheetB', slow('B'));

    await Promise.all([pA, pB]);

    // Both started before either ended — parallel overlap proves
    // per-sheet keying.
    const startA = events.indexOf('start A');
    const startB = events.indexOf('start B');
    const endA = events.indexOf('end A');
    const endB = events.indexOf('end B');
    expect(startA).toBeGreaterThanOrEqual(0);
    expect(startB).toBeGreaterThanOrEqual(0);
    expect(startB).toBeLessThan(endA);  // B started before A ended
    expect(startA).toBeLessThan(endB);  // A started before B ended
  });

  it('returns the awaited result of the queued function', async () => {
    const result = await serializeSheetWrite('Sheet1', async () => 42);
    expect(result).toBe(42);
  });
});
