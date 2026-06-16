import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

/**
 * Unit coverage for the Daily Quest deterministic picker.
 *
 * Quest.js is a Google Apps Script module (not part of the Office src/
 * harness), but questIndexForDate_ is a pure function. We eval just that
 * function out of the file — its body is the only thing exercised, so the
 * GAS-only globals referenced elsewhere in the file are never touched.
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const QUEST_PATH = resolve(__dirname, '..', 'google-apps-script', 'Quest.js');

const code = readFileSync(QUEST_PATH, 'utf8') + '\nreturn { questIndexForDate_ };';
// eslint-disable-next-line no-new-func
const { questIndexForDate_ } = new Function(code)();

describe('questIndexForDate_ (Daily Quest picker)', () => {
  it('is deterministic: same date+salt+n always yields the same index', () => {
    const a = questIndexForDate_('20260615', 'habit', 7);
    const b = questIndexForDate_('20260615', 'habit', 7);
    expect(a).toBe(b);
  });

  it('always returns an index within [0, n)', () => {
    for (let n = 1; n <= 20; n++) {
      for (let day = 1; day <= 28; day++) {
        const dateStr = '202606' + String(day).padStart(2, '0');
        const idx = questIndexForDate_(dateStr, 'task', n);
        expect(idx).toBeGreaterThanOrEqual(0);
        expect(idx).toBeLessThan(n);
      }
    }
  });

  it('returns -1 when there are no candidates', () => {
    expect(questIndexForDate_('20260615', 'habit', 0)).toBe(-1);
    expect(questIndexForDate_('20260615', 'task', -3)).toBe(-1);
  });

  it('uses the salt to make habit and task picks independent streams', () => {
    // Over many days the two salts should not always collapse to the same
    // index (they would if the salt were ignored).
    let differs = 0;
    for (let day = 1; day <= 28; day++) {
      const dateStr = '202606' + String(day).padStart(2, '0');
      if (questIndexForDate_(dateStr, 'habit', 5) !== questIndexForDate_(dateStr, 'task', 5)) {
        differs++;
      }
    }
    expect(differs).toBeGreaterThan(0);
  });

  it('varies the pick across different dates', () => {
    const picks = new Set();
    for (let day = 1; day <= 14; day++) {
      const dateStr = '202606' + String(day).padStart(2, '0');
      picks.add(questIndexForDate_(dateStr, 'habit', 6));
    }
    // Not all 14 days should land on the same habit.
    expect(picks.size).toBeGreaterThan(1);
  });
});
