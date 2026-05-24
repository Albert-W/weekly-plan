import { describe, it, expect } from 'vitest';
import { escapeCSV, formatExcelTime } from './harness.js';

describe('escapeCSV', () => {
  it('returns plain strings unmodified', () => {
    expect(escapeCSV('hello')).toBe('hello');
    expect(escapeCSV('123')).toBe('123');
  });

  it('quotes values containing commas', () => {
    expect(escapeCSV('a,b')).toBe('"a,b"');
  });

  it('quotes and doubles internal quotes', () => {
    expect(escapeCSV('say "hi"')).toBe('"say ""hi"""');
  });

  it('quotes values containing newlines', () => {
    expect(escapeCSV('line1\nline2')).toBe('"line1\nline2"');
  });

  it('returns empty string for null/undefined', () => {
    expect(escapeCSV(null)).toBe('');
    expect(escapeCSV(undefined)).toBe('');
  });

  it('coerces non-string values via String()', () => {
    expect(escapeCSV(42)).toBe('42');
    expect(escapeCSV(0)).toBe('0');
    expect(escapeCSV(true)).toBe('true');
  });

  // ----- CSV formula-injection guard (task #30) -----
  describe('formula-injection guard', () => {
    it('prefixes string starting with = with a single quote', () => {
      expect(escapeCSV('=SUM(A1)')).toBe("'=SUM(A1)");
    });

    it('prefixes string starting with + with a single quote', () => {
      expect(escapeCSV('+HYPERLINK("evil","x")')).toBe('"\'+HYPERLINK(""evil"",""x"")"');
      // The leading quote is added, then the comma/quote rule wraps it.
    });

    it('prefixes string starting with - with a single quote', () => {
      expect(escapeCSV('-1+1')).toBe("'-1+1");
    });

    it('prefixes string starting with @ with a single quote', () => {
      expect(escapeCSV('@cmd')).toBe("'@cmd");
    });

    it('prefixes string starting with tab or carriage return', () => {
      expect(escapeCSV('\tfoo')).toBe("'\tfoo");
      expect(escapeCSV('\rfoo')).toBe("'\rfoo");
    });

    it('does NOT prefix negative NUMBERS (scores like -0.4 must stay numeric)', () => {
      expect(escapeCSV(-0.4)).toBe('-0.4');
      expect(escapeCSV(-1)).toBe('-1');
    });

    it('does NOT prefix strings that only contain a risky char mid-value', () => {
      expect(escapeCSV('a=b')).toBe('a=b');
      expect(escapeCSV('1+2')).toBe('1+2');
    });

    it('combines guard + comma/quote escaping', () => {
      // Has both a leading = AND an embedded comma.
      expect(escapeCSV('=A,B')).toBe('"\'=A,B"');
    });
  });
});

describe('formatExcelTime', () => {
  it('formats fraction-of-day numbers as HH:MM', () => {
    expect(formatExcelTime(0)).toBe('00:00');
    expect(formatExcelTime(0.5)).toBe('12:00');
    expect(formatExcelTime(0.25)).toBe('06:00');
    // 15:30 = 15.5/24 = 0.6458333...
    expect(formatExcelTime(15.5 / 24)).toBe('15:30');
  });

  it('handles whole-number hour values too (e.g. cell stored as 8)', () => {
    // Note: per implementation, any number is multiplied by 24.
    // A literal 8 would render as 8*24 = 192:00 — undesirable but
    // captured here so we notice if the helper is changed.
    expect(formatExcelTime(8)).toBe('192:00');
  });

  it('passes through strings as-is', () => {
    expect(formatExcelTime('08:00')).toBe('08:00');
    expect(formatExcelTime('hello')).toBe('hello');
  });

  it('coerces other types via String()', () => {
    expect(formatExcelTime(null)).toBe('null');
    expect(formatExcelTime(undefined)).toBe('undefined');
  });
});
