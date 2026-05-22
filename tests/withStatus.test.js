import { describe, it, expect, beforeEach, vi } from 'vitest';

/**
 * withStatus wrapper tests (task #6).
 *   - Returns the resolved value on success.
 *   - Catches exceptions, logs them, returns undefined.
 *   - Surfaces failures via showStatus with the supplied label.
 */

const withStatus = globalThis.withStatus;
const showStatus = globalThis.showStatus;

describe('withStatus', () => {
  beforeEach(() => {
    // Replace the status banner with a fresh empty element so we can
    // inspect what was last shown.
    let el = document.getElementById('status');
    if (!el) {
      el = document.createElement('div');
      el.id = 'status';
      document.body.appendChild(el);
    }
    el.textContent = '';
    el.className = 'status';
  });

  it('returns the resolved value when the wrapped function succeeds', async () => {
    const result = await withStatus('do thing', async () => 42);
    expect(result).toBe(42);
  });

  it('returns undefined and shows an error banner when the wrapped function throws', async () => {
    const result = await withStatus('Doing thing', async () => {
      throw new Error('boom');
    });
    expect(result).toBeUndefined();

    const banner = document.getElementById('status');
    expect(banner.textContent).toContain('Doing thing failed');
    expect(banner.textContent).toContain('boom');
    expect(banner.className).toContain('error');
  });

  it('does not show a success banner — that is the wrapped fn\'s job', async () => {
    await withStatus('Quiet op', async () => 'ok');
    const banner = document.getElementById('status');
    expect(banner.textContent).toBe(''); // wrapper added nothing
  });

  it('logs the failure to console.error', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    await withStatus('Log this', async () => {
      throw new Error('nope');
    });
    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });
});
