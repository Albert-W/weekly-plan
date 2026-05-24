import { describe, it, expect } from 'vitest';
import {
  formatDateYYYYMMDD,
  formatDateTime,
  columnLetterToIndex,
  indexToColumnLetter,
  getMonday,
  daysBetween,
  parseAddress,
  getTaskColForDay,
  getScoreColForDay,
  getTaskColLetterForDay,
  getScoreColLetterForDay,
} from './harness.js';

describe('utils.js - date helpers', () => {
  it('formatDateYYYYMMDD pads month and day to 2 digits', () => {
    expect(formatDateYYYYMMDD(new Date(2024, 0, 5))).toBe('20240105');
    expect(formatDateYYYYMMDD(new Date(2024, 11, 31))).toBe('20241231');
  });

  it('formatDateTime appends HH:MM:SS', () => {
    const d = new Date(2024, 0, 5, 9, 7, 3);
    expect(formatDateTime(d)).toBe('20240105 09:07:03');
  });

  it('getMonday returns the same date when given a Monday', () => {
    const mon = new Date(2024, 0, 1); // Mon Jan 1 2024
    const result = getMonday(mon);
    expect(result.getDay()).toBe(1);
    expect(result.getDate()).toBe(1);
  });

  it('getMonday returns the previous Monday when given a Sunday', () => {
    const sun = new Date(2024, 0, 7); // Sun Jan 7 2024
    const result = getMonday(sun);
    expect(result.getDay()).toBe(1);
    expect(result.getDate()).toBe(1);
  });

  it('getMonday zeroes the time component', () => {
    const wed = new Date(2024, 0, 3, 15, 42, 11);
    const result = getMonday(wed);
    expect(result.getHours()).toBe(0);
    expect(result.getMinutes()).toBe(0);
    expect(result.getSeconds()).toBe(0);
    expect(result.getMilliseconds()).toBe(0);
  });

  it('daysBetween counts whole days', () => {
    const a = new Date(2024, 0, 1);
    const b = new Date(2024, 0, 8);
    expect(daysBetween(a, b)).toBe(7);
    expect(daysBetween(b, a)).toBe(-7);
  });
});

describe('utils.js - column math', () => {
  it('columnLetterToIndex returns 0-based index for single letters', () => {
    expect(columnLetterToIndex('A')).toBe(0);
    expect(columnLetterToIndex('B')).toBe(1);
    expect(columnLetterToIndex('Q')).toBe(16);
    expect(columnLetterToIndex('Z')).toBe(25);
  });

  it('columnLetterToIndex returns 0-based index for multi-letter columns', () => {
    expect(columnLetterToIndex('AA')).toBe(26);
    expect(columnLetterToIndex('AB')).toBe(27);
    expect(columnLetterToIndex('AZ')).toBe(51);
    expect(columnLetterToIndex('BA')).toBe(52);
    expect(columnLetterToIndex('ZZ')).toBe(701);
  });

  it('indexToColumnLetter is the inverse for single-letter columns', () => {
    for (const letter of ['A', 'C', 'Q', 'Z']) {
      expect(indexToColumnLetter(columnLetterToIndex(letter))).toBe(letter);
    }
  });

  it('indexToColumnLetter handles multi-letter columns correctly', () => {
    expect(indexToColumnLetter(26)).toBe('AA');
    expect(indexToColumnLetter(27)).toBe('AB');
    expect(indexToColumnLetter(51)).toBe('AZ');
  });

  it('round-trips through both functions for every 1- and 2-letter column', () => {
    // Verify the inverse property end-to-end. Drops the previous
    // "broken on purpose" pin (#39).
    for (let i = 0; i < 26 * 27; i++) {
      const letter = indexToColumnLetter(i);
      expect(columnLetterToIndex(letter)).toBe(i);
    }
  });
});

describe('utils.js - parseAddress', () => {
  it('parses single-cell addresses', () => {
    expect(parseAddress('A1')).toEqual({ column: 'A', colIndex: 1, row: 1 });
    expect(parseAddress('Q38')).toEqual({ column: 'Q', colIndex: 17, row: 38 });
  });

  it('returns null on garbage', () => {
    expect(parseAddress('not-an-address')).toBeNull();
  });
});

describe('utils.js - day/column helpers', () => {
  it('getTaskColForDay returns the expected 1-based columns 3,5,7,9,11,13,15', () => {
    expect([0, 1, 2, 3, 4, 5, 6].map(getTaskColForDay)).toEqual([3, 5, 7, 9, 11, 13, 15]);
  });

  it('getScoreColForDay returns the expected 1-based columns 4,6,8,10,12,14,16', () => {
    expect([0, 1, 2, 3, 4, 5, 6].map(getScoreColForDay)).toEqual([4, 6, 8, 10, 12, 14, 16]);
  });

  it('getTaskColLetterForDay returns C,E,G,I,K,M,O', () => {
    expect([0, 1, 2, 3, 4, 5, 6].map(getTaskColLetterForDay)).toEqual(
      ['C', 'E', 'G', 'I', 'K', 'M', 'O']
    );
  });

  it('getScoreColLetterForDay returns D,F,H,J,L,N,P', () => {
    expect([0, 1, 2, 3, 4, 5, 6].map(getScoreColLetterForDay)).toEqual(
      ['D', 'F', 'H', 'J', 'L', 'N', 'P']
    );
  });
});
