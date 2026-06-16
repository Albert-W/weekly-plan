import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

/**
 * Unit coverage for the Quest Streak Combo multiplier curve.
 *
 * Quest.js is a Google Apps Script module, but comboMultiplier_ is a pure
 * function. We eval just that function out of the file (its body is the
 * only thing exercised, so the DocumentProperties/LockService globals
 * referenced elsewhere are never touched).
 *
 * Curve: multiplier(n) = BASE + min(n, CAP)*STEP. With BASE=1.0, STEP=0.2,
 * CAP=5 => 0d x1.0, 1d x1.2, 2d x1.4, 3d x1.6, 4d x1.8, 5d+ x2.0.
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const QUEST_PATH = resolve(__dirname, '..', 'google-apps-script', 'Quest.js');

const code = readFileSync(QUEST_PATH, 'utf8') + '\nreturn { comboMultiplier_ };';
// eslint-disable-next-line no-new-func
const { comboMultiplier_ } = new Function(code)();

const BASE = 1.0;
const STEP = 0.2;
const CAP = 5;
const m = (n) => comboMultiplier_(n, BASE, STEP, CAP);
const near = (a, b) => Math.abs(a - b) < 1e-9;

describe('comboMultiplier_ (quest streak combo curve)', () => {
  it('returns BASE (no bonus) for 0 or negative days', () => {
    expect(m(0)).toBe(1.0);
    expect(m(-3)).toBe(1.0);
  });

  it('rises 0.2 per consecutive day', () => {
    expect(near(m(1), 1.2)).toBe(true);
    expect(near(m(2), 1.4)).toBe(true);
    expect(near(m(3), 1.6)).toBe(true);
    expect(near(m(4), 1.8)).toBe(true);
    expect(near(m(5), 2.0)).toBe(true);
  });

  it('caps at CAP days', () => {
    expect(near(m(6), 2.0)).toBe(true);
    expect(near(m(50), 2.0)).toBe(true);
  });

  it('matches the documented example (3-day combo on a 1.5pt quest task)', () => {
    // featured task 1.5 pts * combo x1.6 = 2.4
    expect(near(1.5 * m(3), 2.4)).toBe(true);
  });
});
