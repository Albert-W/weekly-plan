import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

/**
 * Unit coverage for the XP level curve.
 *
 * Xp.js is a Google Apps Script module, but levelForXp_ is a pure
 * function. We eval just that function out of the file (its body is the
 * only thing exercised, so the DocumentProperties/LockService globals
 * referenced elsewhere are never touched).
 *
 * Curve: XP to advance level L is BASE + (L-1)*STEP. With BASE=50, STEP=25
 * cumulative thresholds are L1=0, L2=50, L3=125, L4=225, L5=350.
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const XP_PATH = resolve(__dirname, '..', 'google-apps-script', 'Xp.js');

const code = readFileSync(XP_PATH, 'utf8') + '\nreturn { levelForXp_ };';
// eslint-disable-next-line no-new-func
const { levelForXp_ } = new Function(code)();

const BASE = 50;
const STEP = 25;

describe('levelForXp_ (XP level curve)', () => {
  it('starts at level 1 with 0 XP', () => {
    const r = levelForXp_(0, BASE, STEP);
    expect(r.level).toBe(1);
    expect(r.levelFloor).toBe(0);
    expect(r.nextThreshold).toBe(50);
  });

  it('treats negative/invalid XP as 0 (level 1)', () => {
    expect(levelForXp_(-10, BASE, STEP).level).toBe(1);
  });

  it('levels up exactly at each cumulative threshold', () => {
    expect(levelForXp_(49, BASE, STEP).level).toBe(1);
    expect(levelForXp_(50, BASE, STEP).level).toBe(2); // 50
    expect(levelForXp_(124, BASE, STEP).level).toBe(2);
    expect(levelForXp_(125, BASE, STEP).level).toBe(3); // 50+75
    expect(levelForXp_(225, BASE, STEP).level).toBe(4); // +100
    expect(levelForXp_(350, BASE, STEP).level).toBe(5); // +125
  });

  it('reports the current level floor and next threshold', () => {
    const r = levelForXp_(200, BASE, STEP); // within level 3 [125, 225)
    expect(r.level).toBe(3);
    expect(r.levelFloor).toBe(125);
    expect(r.nextThreshold).toBe(225);
  });

  it('is monotonic: level never decreases as XP grows', () => {
    let prev = 1;
    for (let xp = 0; xp <= 1000; xp += 7) {
      const lvl = levelForXp_(xp, BASE, STEP).level;
      expect(lvl).toBeGreaterThanOrEqual(prev);
      prev = lvl;
    }
  });
});
