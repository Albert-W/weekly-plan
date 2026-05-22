/**
 * In-memory fake of the Office.js Excel object model.
 *
 * Covers ONLY what the production code in src/taskpane/js/ actually
 * uses. This keeps the fake small enough to reason about while
 * faithfully simulating the queue-and-sync semantics that make
 * Office.js different from a typical synchronous spreadsheet API.
 *
 * Public API:
 *   - makeFakeExcel({ sheets, activeSheet })
 *       -> { workbook, runWithContext, installAsExcelGlobal, helpers }
 *
 * Each sheet is a Map keyed by cell address ("A1", "B12", ...) with
 * values of shape { value, fillColor }. Cells that have never been
 * touched are treated as undefined / empty by getters.
 *
 * Range objects parse single cells ("A1"), ranges ("A1:B5"), and
 * whole columns ("B:B" -> resolved via getUsedRange()).
 *
 * Reads are buffered: a Range.load('values') call only records intent;
 * Range.values is populated when context.sync() runs.
 *
 * Writes (Range.values = ..., format.fill.color = ..., clear(),
 * format.fill.clear()) are queued and applied at the next sync().
 *
 * The sync counter is exposed so tests can assert performance
 * regressions (e.g. "this function MUST stay <=2 syncs").
 */

// ---------------------------------------------------------------------
// Cell address parsing
// ---------------------------------------------------------------------

function colLetterToIndex(letter) {
  let idx = 0;
  for (let i = 0; i < letter.length; i++) {
    idx = idx * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return idx; // 1-based
}

function indexToColLetter(index) {
  // index is 1-based
  let letter = '';
  while (index > 0) {
    const rem = (index - 1) % 26;
    letter = String.fromCharCode('A'.charCodeAt(0) + rem) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function cellAddr(col, row) {
  return `${indexToColLetter(col)}${row}`;
}

/**
 * Parse a single cell or range address.
 * Returns { startCol, startRow, endCol, endRow, isWholeColumn }.
 * Whole-column addresses ("B:B") return startRow/endRow as null.
 */
function parseRange(addr) {
  // Whole column like "B:B" or "C:E"
  const wholeColMatch = addr.match(/^([A-Z]+):([A-Z]+)$/);
  if (wholeColMatch) {
    return {
      startCol: colLetterToIndex(wholeColMatch[1]),
      startRow: null,
      endCol: colLetterToIndex(wholeColMatch[2]),
      endRow: null,
      isWholeColumn: true,
    };
  }
  // Range like "A1:C5"
  const rangeMatch = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (rangeMatch) {
    return {
      startCol: colLetterToIndex(rangeMatch[1]),
      startRow: parseInt(rangeMatch[2], 10),
      endCol: colLetterToIndex(rangeMatch[3]),
      endRow: parseInt(rangeMatch[4], 10),
      isWholeColumn: false,
    };
  }
  // Single cell "A1"
  const singleMatch = addr.match(/^([A-Z]+)(\d+)$/);
  if (singleMatch) {
    const col = colLetterToIndex(singleMatch[1]);
    const row = parseInt(singleMatch[2], 10);
    return {
      startCol: col,
      startRow: row,
      endCol: col,
      endRow: row,
      isWholeColumn: false,
    };
  }
  throw new Error(`fake-excel: unsupported address "${addr}"`);
}

// ---------------------------------------------------------------------
// Sheet model
// ---------------------------------------------------------------------

function createSheet(name) {
  return {
    name,
    cells: new Map(), // addr -> { value, fillColor }
    /** EventResult records, keyed by event name. */
    handlers: {
      onChanged: [],
      onSelectionChanged: [],
    },
  };
}

function readCell(sheet, addr) {
  return sheet.cells.get(addr) || { value: null, fillColor: null };
}

function writeCell(sheet, addr, patch) {
  const existing = sheet.cells.get(addr) || { value: null, fillColor: null };
  sheet.cells.set(addr, { ...existing, ...patch });
}

/** Find the used extent of a sheet (1-based [minRow, maxRow, minCol, maxCol]). */
function findUsedExtent(sheet, restrictToCols = null) {
  let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
  for (const addr of sheet.cells.keys()) {
    const { startCol, startRow } = parseRange(addr);
    if (restrictToCols) {
      if (startCol < restrictToCols.start || startCol > restrictToCols.end) continue;
    }
    if (startRow < minRow) minRow = startRow;
    if (startRow > maxRow) maxRow = startRow;
    if (startCol < minCol) minCol = startCol;
    if (startCol > maxCol) maxCol = startCol;
  }
  if (maxRow === 0) {
    return { minRow: 1, maxRow: 0, minCol: 1, maxCol: 0, isEmpty: true };
  }
  return { minRow, maxRow, minCol, maxCol, isEmpty: false };
}

// ---------------------------------------------------------------------
// EventResult (returned by .onChanged.add / .onSelectionChanged.add)
// ---------------------------------------------------------------------

function createEventResult(sheet, eventName, callback) {
  const result = {
    id: Symbol('eventResult'),
    callback,
    removed: false,
    remove() {
      result.removed = true;
      const list = sheet.handlers[eventName];
      const i = list.indexOf(result);
      if (i >= 0) list.splice(i, 1);
    },
  };
  sheet.handlers[eventName].push(result);
  return result;
}

// ---------------------------------------------------------------------
// Range
// ---------------------------------------------------------------------

function createRange(sheet, addr, ctx, parsedOverride = null) {
  const parsed = parsedOverride || parseRange(addr);
  const range = {
    _sheet: sheet,
    _addr: addr,
    _parsed: parsed,
    _loaded: new Set(),
    values: undefined,
    rowCount: undefined,
    format: null, // populated lazily below
  };

  range.load = function load(propsArg) {
    const props = Array.isArray(propsArg) ? propsArg : String(propsArg).split('/');
    for (const p of props) range._loaded.add(p);
    ctx._pendingReads.push(range);
    return range;
  };

  range.clear = function clear(_applyTo) {
    ctx._pendingWrites.push({ type: 'clearContents', range });
  };

  range.getOffsetRange = function getOffsetRange(rowOffset, colOffset) {
    const newParsed = {
      startCol: parsed.startCol + colOffset,
      startRow: parsed.startRow + rowOffset,
      endCol: parsed.endCol + colOffset,
      endRow: parsed.endRow + rowOffset,
      isWholeColumn: false,
    };
    const newAddr = newParsed.startCol === newParsed.endCol && newParsed.startRow === newParsed.endRow
      ? cellAddr(newParsed.startCol, newParsed.startRow)
      : `${cellAddr(newParsed.startCol, newParsed.startRow)}:${cellAddr(newParsed.endCol, newParsed.endRow)}`;
    return createRange(sheet, newAddr, ctx, newParsed);
  };

  range.getUsedRange = function getUsedRange() {
    const restrict = parsed.isWholeColumn
      ? { start: parsed.startCol, end: parsed.endCol }
      : null;
    const ext = findUsedExtent(sheet, restrict);
    if (ext.isEmpty) {
      const empty = createRange(sheet, addr, ctx, {
        startCol: parsed.startCol,
        startRow: 1,
        endCol: parsed.endCol,
        endRow: 1,
        isWholeColumn: false,
      });
      empty._isEmptyUsed = true;
      return empty;
    }
    const startCol = parsed.isWholeColumn ? parsed.startCol : ext.minCol;
    const endCol = parsed.isWholeColumn ? parsed.endCol : ext.maxCol;
    const newAddr = `${cellAddr(startCol, ext.minRow)}:${cellAddr(endCol, ext.maxRow)}`;
    return createRange(sheet, newAddr, ctx, {
      startCol,
      startRow: ext.minRow,
      endCol,
      endRow: ext.maxRow,
      isWholeColumn: false,
    });
  };

  // Value getter via property descriptor so writes to `.values = [[x]]`
  // can be intercepted as queued writes.
  Object.defineProperty(range, '_writeValues', {
    value: null,
    writable: true,
  });

  // Format proxy — fill.color setter queues a write.
  const fillProxy = {
    set color(c) {
      ctx._pendingWrites.push({ type: 'fillColor', range, color: c });
    },
    clear() {
      ctx._pendingWrites.push({ type: 'fillClear', range });
    },
  };
  range.format = { fill: fillProxy };

  return range;
}

// Patch in a values setter that queues a write.
function attachValuesSetter(range, ctx) {
  let backingValues;
  range._setBackingValues = (v) => { backingValues = v; };
  Object.defineProperty(range, 'values', {
    configurable: true,
    get() { return backingValues; },
    set(v) {
      backingValues = v;
      ctx._pendingWrites.push({ type: 'values', range, values: v });
    },
  });
}

// ---------------------------------------------------------------------
// Worksheets collection
// ---------------------------------------------------------------------

function createWorksheets(workbook, ctx) {
  return {
    _loadedItems: false,
    items: [], // populated on sync after a load('items/name')

    getItem(name) {
      const s = workbook.sheets.get(name);
      if (!s) throw new Error(`fake-excel: sheet "${name}" not found`);
      return createSheetProxy(s, ctx);
    },

    getItemOrNullObject(name) {
      const s = workbook.sheets.get(name);
      if (s) return createSheetProxy(s, ctx);
      return createNullSheetProxy(name, ctx);
    },

    getActiveWorksheet() {
      const s = workbook.sheets.get(workbook.activeSheet);
      if (!s) throw new Error(`fake-excel: no active sheet`);
      return createSheetProxy(s, ctx);
    },

    load(prop) {
      // We only care that this was requested; items get filled at sync.
      ctx._worksheetsLoad = prop;
      return this;
    },

    onActivated: {
      add(callback) {
        workbook.onActivatedHandlers.push({
          callback,
          remove() {
            const i = workbook.onActivatedHandlers.findIndex((h) => h.callback === callback);
            if (i >= 0) workbook.onActivatedHandlers.splice(i, 1);
          },
        });
        return workbook.onActivatedHandlers[workbook.onActivatedHandlers.length - 1];
      },
    },
  };
}

function createSheetProxy(sheet, ctx) {
  const proxy = {
    _sheet: sheet,
    isNullObject: false,
    name: undefined,
    getRange(addr) {
      return createRangeAndAttach(sheet, addr, ctx);
    },
    getUsedRange() {
      return createRangeAndAttach(sheet, 'A:Z', ctx).getUsedRange();
    },
    activate() {
      ctx._activatedSheetName = sheet.name;
    },
    load(prop) {
      ctx._pendingSheetNameLoads.push({ proxy, prop });
      return proxy;
    },
    onChanged: {
      add(callback) {
        return createEventResult(sheet, 'onChanged', callback);
      },
    },
    onSelectionChanged: {
      add(callback) {
        return createEventResult(sheet, 'onSelectionChanged', callback);
      },
    },
  };
  return proxy;
}

function createNullSheetProxy(name, ctx) {
  return {
    _sheet: null,
    isNullObject: true,
    name,
    getRange() { throw new Error(`fake-excel: getRange on null sheet "${name}"`); },
  };
}

function createRangeAndAttach(sheet, addr, ctx) {
  const range = createRange(sheet, addr, ctx);
  attachValuesSetter(range, ctx);
  return range;
}

// ---------------------------------------------------------------------
// Context (the thing the production code calls "context")
// ---------------------------------------------------------------------

function createContext(workbook) {
  const ctx = {
    _pendingReads: [],
    _pendingWrites: [],
    _pendingSheetNameLoads: [],
    _worksheetsLoad: null,
    _activatedSheetName: null,
    _syncCount: 0,
  };

  ctx.workbook = { worksheets: createWorksheets(workbook, ctx) };

  ctx.sync = async function sync() {
    ctx._syncCount++;

    // 1) Apply queued writes first (mirrors Office.js: writes within a
    //    sync flush before reads are evaluated, but practically all
    //    consumers either read OR write in a single sync).
    for (const w of ctx._pendingWrites) {
      const r = w.range;
      const startCol = r._parsed.startCol;
      const startRow = r._parsed.startRow;
      const endCol = r._parsed.endCol;
      const endRow = r._parsed.endRow;

      if (w.type === 'values') {
        const rows = endRow - startRow + 1;
        const cols = endCol - startCol + 1;
        for (let i = 0; i < rows; i++) {
          for (let j = 0; j < cols; j++) {
            const v = w.values[i] !== undefined ? w.values[i][j] : undefined;
            writeCell(r._sheet, cellAddr(startCol + j, startRow + i), { value: v });
          }
        }
      } else if (w.type === 'fillColor') {
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            writeCell(r._sheet, cellAddr(col, row), { fillColor: w.color });
          }
        }
      } else if (w.type === 'fillClear') {
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            writeCell(r._sheet, cellAddr(col, row), { fillColor: null });
          }
        }
      } else if (w.type === 'clearContents') {
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            writeCell(r._sheet, cellAddr(col, row), { value: null });
          }
        }
      }
    }
    ctx._pendingWrites.length = 0;

    // 2) Populate values on loaded ranges
    for (const r of ctx._pendingReads) {
      if (r._loaded.has('values')) {
        const rows = (r._parsed.endRow ?? r._parsed.startRow) - r._parsed.startRow + 1;
        const cols = r._parsed.endCol - r._parsed.startCol + 1;
        const out = [];
        for (let i = 0; i < rows; i++) {
          const rowArr = [];
          for (let j = 0; j < cols; j++) {
            const c = readCell(r._sheet, cellAddr(r._parsed.startCol + j, r._parsed.startRow + i));
            rowArr.push(c.value === null ? '' : c.value);
          }
          out.push(rowArr);
        }
        // Refresh the setter's backing value WITHOUT removing the
        // setter — otherwise subsequent `range.values = ...` writes
        // would silently bypass the queue.
        r._setBackingValues(out);
      }
      if (r._loaded.has('rowCount')) {
        const rows = r._parsed.endRow !== null
          ? r._parsed.endRow - r._parsed.startRow + 1
          : 0;
        r.rowCount = rows;
      }
    }
    ctx._pendingReads.length = 0;

    // 3) Populate sheet name proxies
    for (const { proxy, prop } of ctx._pendingSheetNameLoads) {
      if (prop === 'name') proxy.name = proxy._sheet.name;
    }
    ctx._pendingSheetNameLoads.length = 0;

    // 4) Populate worksheets.items if requested
    if (ctx._worksheetsLoad === 'items/name') {
      ctx.workbook.worksheets.items = Array.from(workbook.sheets.values()).map((s) => ({
        name: s.name,
      }));
      ctx._worksheetsLoad = null;
    }

    // 5) Apply activate() if any
    if (ctx._activatedSheetName) {
      workbook.activeSheet = ctx._activatedSheetName;
      ctx._activatedSheetName = null;
    }
  };

  return ctx;
}

// ---------------------------------------------------------------------
// Public factory
// ---------------------------------------------------------------------

/**
 * @param {Object} [opts]
 * @param {string[]} [opts.sheets] - sheet names to create (defaults to none)
 * @param {string} [opts.activeSheet] - which sheet is active (defaults to first)
 */
export function makeFakeExcel(opts = {}) {
  const sheetNames = opts.sheets || [];
  const workbook = {
    sheets: new Map(sheetNames.map((n) => [n, createSheet(n)])),
    activeSheet: opts.activeSheet || sheetNames[0] || null,
    onActivatedHandlers: [],
  };

  const totalSyncCounter = { count: 0 };

  async function runWithContext(fn) {
    const ctx = createContext(workbook);
    const result = await fn(ctx);
    totalSyncCounter.count += ctx._syncCount;
    return { result, syncCount: ctx._syncCount };
  }

  // Install a stub of the Excel global so production code can call
  // Excel.run(fn) directly.
  function installAsExcelGlobal(globalObj = globalThis) {
    globalObj.Excel = {
      run: async (fn) => {
        const ctx = createContext(workbook);
        const ret = await fn(ctx);
        totalSyncCounter.count += ctx._syncCount;
        return ret;
      },
      ClearApplyTo: { contents: 'contents' },
    };
  }

  const helpers = {
    /** Bulk-set cells: setCells('Sheet', { 'A1': 'foo', 'B2': 3 }). */
    setCells(sheetName, map) {
      const sheet = workbook.sheets.get(sheetName);
      for (const [addr, value] of Object.entries(map)) {
        writeCell(sheet, addr, { value });
      }
    },
    setFill(sheetName, addr, color) {
      const sheet = workbook.sheets.get(sheetName);
      writeCell(sheet, addr, { fillColor: color });
    },
    getCellValue(sheetName, addr) {
      const sheet = workbook.sheets.get(sheetName);
      const c = readCell(sheet, addr);
      return c.value;
    },
    getCellColor(sheetName, addr) {
      const sheet = workbook.sheets.get(sheetName);
      const c = readCell(sheet, addr);
      return c.fillColor;
    },
    getSyncCount() {
      return totalSyncCounter.count;
    },
    resetSyncCount() {
      totalSyncCounter.count = 0;
    },
    /** Manually trigger a sheet event handler (for testing handler de-dup). */
    triggerSheetEvent(sheetName, eventName, eventObj) {
      const sheet = workbook.sheets.get(sheetName);
      const handlers = sheet.handlers[eventName] || [];
      return Promise.all(handlers.map((h) => h.callback(eventObj)));
    },
    /** Inspect registered handlers for assertions. */
    getHandlerCount(sheetName, eventName) {
      const sheet = workbook.sheets.get(sheetName);
      return (sheet.handlers[eventName] || []).length;
    },
  };

  return { workbook, runWithContext, installAsExcelGlobal, helpers };
}
