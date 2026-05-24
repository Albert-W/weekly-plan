/**
 * Per-sheet write serialization for the Combined Tracker Add-in.
 *
 * Problem
 * -------
 * Several handler functions implement a read-modify-write against the
 * same cell (e.g. processWeeklyScoreChange reads the daily total,
 * adds a delta, writes the new total back). Office.js fires the
 * onChanged event per cell write, and may dispatch multiple async
 * handlers concurrently when the user pastes a column of values or
 * types quickly. Two overlapping handlers can each read the same
 * "current total", each compute "current + delta", and each write
 * back — losing one increment.
 *
 * Solution
 * --------
 * Per-sheet Promise chain. The dispatch site (events.js) wraps the
 * handler call in `serializeSheetWrite(sheetName, () => handler(...))`.
 * All work targeting the same sheet name runs sequentially regardless
 * of how many concurrent dispatches Office.js fires; work targeting
 * different sheets runs independently in parallel.
 *
 * The chain swallows rejections so one thrown handler can't poison
 * future writes; the handler is responsible for its own error
 * surfacing (typically via withStatus).
 *
 * Test-only helpers (resetWriteChains, getInFlightChain) live here
 * too so tests can deterministically observe the serialization.
 */

const _writeChains = new Map();

/**
 * Queue an async function to run after all previously-queued work
 * for the same sheet name has completed. Returns the promise the
 * caller can await.
 *
 * @param {string} sheetName - Key for the per-sheet chain.
 * @param {() => Promise<any>} fn - Async work to serialize.
 * @returns {Promise<any>}
 */
function serializeSheetWrite(sheetName, fn) {
  const prev = _writeChains.get(sheetName) || Promise.resolve();
  // .catch(() => {}) keeps the chain alive after a thrown handler.
  // The handler itself is expected to surface its own errors.
  const next = prev.catch(() => {}).then(fn);
  _writeChains.set(sheetName, next);
  return next;
}

/**
 * Reset all in-flight chains. Test-only — production code never
 * needs this because the chains naturally drain as their work
 * resolves.
 */
function resetWriteChains() {
  _writeChains.clear();
}

/**
 * Return the current in-flight Promise for a sheet, or a resolved
 * promise if nothing is queued. Test-only.
 *
 * @param {string} sheetName
 * @returns {Promise<any>}
 */
function getInFlightChain(sheetName) {
  return _writeChains.get(sheetName) || Promise.resolve();
}

window.serializeSheetWrite = serializeSheetWrite;
window.resetWriteChains = resetWriteChains;
window.getInFlightChain = getInFlightChain;
