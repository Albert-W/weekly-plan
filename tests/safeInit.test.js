import { describe, it, expect, vi } from 'vitest';
import { safeInit } from './harness.js';

/**
 * safeInit unit tests (task #35).
 *
 * Contract:
 *  - On success, returns the awaited fn() result.
 *  - On failure, logs `${label}: ${e.message}` and returns null —
 *    does NOT propagate the error.
 *  - Tolerates a thrown non-Error value.
 */

describe('safeInit', () => {
  it('returns the awaited fn() result on success', async () => {
    const result = await safeInit('label', async () => 42);
    expect(result).toBe(42);
  });

  it('returns null and logs on failure', async () => {
    const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});
    const result = await safeInit('Habits init skipped', async () => {
      throw new Error('boom');
    });
    expect(result).toBeNull();
    expect(logSpy).toHaveBeenCalledWith('Habits init skipped:', 'boom');
    logSpy.mockRestore();
  });

  it('tolerates a thrown non-Error (string) value', async () => {
    const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});
    const result = await safeInit('weird', async () => {
      throw 'string-error';
    });
    expect(result).toBeNull();
    expect(logSpy).toHaveBeenCalledWith('weird:', 'string-error');
    logSpy.mockRestore();
  });

  it('does NOT propagate the error to the caller', async () => {
    // The whole point: caller should never need try/catch around safeInit.
    await expect(
      safeInit('x', async () => { throw new Error('nope'); })
    ).resolves.toBeNull();
  });
});
