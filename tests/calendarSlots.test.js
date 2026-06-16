import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

/**
 * Unit coverage for the Calendar→Weekly slot-mapping math.
 *
 * Calendar.js is a Google Apps Script module, but eventSlotIndices_ is a
 * pure function. We eval just that function out of the file (its body is
 * the only thing exercised, so the CalendarApp/Spreadsheet globals
 * referenced elsewhere are never touched).
 *
 * Grid model: 30-min rows starting at FIRST_HOUR=8, slotCount=32
 * (rows 5..36 => 8:00 .. 23:30). Slot i covers [8 + i*0.5, 8 + i*0.5 + 0.5).
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const CAL_PATH = resolve(__dirname, '..', 'google-apps-script', 'Calendar.js');

const code = readFileSync(CAL_PATH, 'utf8') + '\nreturn { eventSlotIndices_ };';
// eslint-disable-next-line no-new-func
const { eventSlotIndices_ } = new Function(code)();

const FIRST = 8;
const SLOTS = 32;

describe('eventSlotIndices_ (Calendar → Weekly mapping)', () => {
  it('maps an 8:00–8:30 event to the first slot only', () => {
    expect(eventSlotIndices_(8, 8.5, FIRST, SLOTS)).toEqual([0]);
  });

  it('maps a 9:00–10:30 event to three consecutive slots', () => {
    // 9:00 -> index 2, 9:30 -> 3, 10:00 -> 4 ; 10:30 is exclusive end
    expect(eventSlotIndices_(9, 10.5, FIRST, SLOTS)).toEqual([2, 3, 4]);
  });

  it('maps the final 23:30 slot correctly', () => {
    // last slot index 31 covers [23.5, 24)
    expect(eventSlotIndices_(23.5, 24, FIRST, SLOTS)).toEqual([31]);
  });

  it('covers a slot when the event partially overlaps it', () => {
    // 9:15–9:45 overlaps slot 2 (9:00–9:30) and slot 3 (9:30–10:00)
    expect(eventSlotIndices_(9.25, 9.75, FIRST, SLOTS)).toEqual([2, 3]);
  });

  it('returns no slots for a span entirely outside the grid window', () => {
    expect(eventSlotIndices_(6, 7, FIRST, SLOTS)).toEqual([]); // before 8:00
    expect(eventSlotIndices_(24, 25, FIRST, SLOTS)).toEqual([]); // at/after 24:00
  });

  it('never returns an index outside [0, slotCount)', () => {
    const idx = eventSlotIndices_(8, 24, FIRST, SLOTS);
    expect(idx[0]).toBe(0);
    expect(idx[idx.length - 1]).toBe(SLOTS - 1);
    expect(idx.length).toBe(SLOTS);
  });
});
