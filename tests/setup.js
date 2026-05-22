/**
 * Vitest setup file. Runs once before each test file.
 *
 * Loads the production src/ JavaScript modules via fs+eval (mirrors
 * how taskpane.html loads them via <script> tags) so tests can call
 * the production functions directly via window.xxx globals.
 *
 * Office.js / Excel globals are stubbed first; the real Excel global
 * is replaced per-test via fakeExcel.installAsExcelGlobal() when
 * needed.
 */

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installOfficeGlobal } from './mocks/office.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const SRC_DIR = resolve(__dirname, '..', 'src', 'taskpane', 'js');

// Load order mirrors taskpane.html (config -> state -> utils -> registry ->
// ui -> tasks -> summary -> export -> habits -> weekly -> events).
const SRC_FILES = [
  'config.js',
  'state.js',
  'utils.js',
  'registry.js',
  'ui.js',
  'tasks.js',
  'summary.js',
  'export.js',
  'habits.js',
  'weekly.js',
  'events.js',
  // intentionally NOT loading app.js — it only contains Office.onReady
  // bootstrapping that we don't want firing in tests.
];

// 1. Install Office stub before any src file runs (config/state are
//    fine without it, but ui.js references Office.context inside
//    functions; loading the file is safe because functions aren't
//    invoked here).
installOfficeGlobal(globalThis);

// 2. Some src files end with `window.foo = foo;` exports. In jsdom,
//    window === globalThis, but the `const` declarations at the top
//    of each file would normally be scoped to the script. We wrap
//    each src file's body so it runs in a function scope but the
//    `window.foo = foo;` assignments still attach to the global.
//
//    We also need to neutralize the `const CONFIG = {...}; ...
//    window.CONFIG = CONFIG;` pattern: re-running the file would
//    re-declare and re-assign, which is fine.
for (const file of SRC_FILES) {
  const path = resolve(SRC_DIR, file);
  const code = readFileSync(path, 'utf8');
  // eslint-disable-next-line no-new-func
  new Function(code).call(globalThis);
}

// 3. Sanity check: at least one expected global should now exist.
if (typeof globalThis.CONFIG === 'undefined') {
  throw new Error('Test setup failed: CONFIG global was not attached.');
}
