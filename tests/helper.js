/* ============================================================
 * Test harness — loads the browser app (i18n.js + vendored SheetJS +
 * app.js) into a Node VM with DOM stubs, then exposes the pure
 * calculation/parse/format functions via `__app` so tests can call them
 * exactly as the page does. app.js is upload-driven and does no DOM work
 * at load time, so nothing renders here — we only exercise the logic.
 *
 * Why a VM instead of import: app.js is a classic browser <script>
 * (globals, no exports). We concatenate the three files and append a
 * single export line that captures the lexical consts (short/full/…,
 * LANG) alongside the function declarations, all in one script scope.
 * ============================================================ */
'use strict';
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const ASSETS = path.join(__dirname, '..', 'assets');

// Build a fresh VM with the app loaded. `lang` seeds localStorage so the
// locale-dependent formatters (short/full/pct1…) can be tested per language.
function makeCtx(lang = 'es') {
  // a permissive DOM element stub — never actually used for rendering here
  const el = {
    style: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    setAttribute() {}, addEventListener() {}, appendChild() {}, removeChild() {},
    closest() { return null; }, contains() { return false; }, querySelectorAll() { return []; },
    scrollIntoView() {}, focus() {}, select() {}, click() {}, hidden: false,
    textContent: '', innerHTML: '', value: ''
  };
  const ctx = {
    console, Date, Math, Number, String, Array, Object, JSON, parseFloat, parseInt,
    isFinite, isNaN, Boolean, Map, Set, RegExp, Error, Promise, Uint8Array, ArrayBuffer, Infinity, NaN,
    Blob, Response, CompressionStream, DecompressionStream, TextEncoder, TextDecoder, btoa, atob,
    setTimeout, clearTimeout, requestAnimationFrame: cb => setTimeout(cb, 0),
    localStorage: { s: { 'mro.lang': lang }, getItem(k) { return this.s[k] || null; }, setItem(k, v) { this.s[k] = v; } },
    getComputedStyle: () => ({ getPropertyValue: () => '#000000' }),
    navigator: { clipboard: { writeText: () => Promise.resolve() } },
    location: { origin: 'https://example.test', pathname: '/', hash: '', search: '' },
    history: { replaceState() {} },
    document: {
      getElementById: () => el, querySelector: () => el, querySelectorAll: () => [],
      addEventListener() {}, createElement: () => el, body: el,
      documentElement: { setAttribute() {}, getAttribute() { return lang; } }
    }
  };
  ctx.window = ctx; ctx.globalThis = ctx; ctx.self = ctx;
  vm.createContext(ctx);

  const src = ['i18n.js', 'vendor/xlsx.full.min.js', 'app.js']
    .map(f => fs.readFileSync(path.join(ASSETS, f), 'utf8'))
    .join('\n;\n')
    // capture lexical consts + function declarations from app.js's scope
    + `\n;globalThis.__app = {
         cleanData, computeModel, buildSampleWorkbook, compactModel, hydrateModel,
         toISO, lyISO, short, full, num, pct1, pct0, fmt, encodeShare, decodeShare,
         setLang: l => { LANG = l; }, getLang: () => LANG
       };`;
  vm.runInContext(src, ctx, { filename: 'app-bundle.js' });
  return ctx;
}

// Serialize sheets ({name: aoa}) into an .xlsx byte array via SheetJS.
function sheetsToBytes(ctx, sheets) {
  const wb = ctx.XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) {
    ctx.XLSX.utils.book_append_sheet(wb, ctx.XLSX.utils.aoa_to_sheet(aoa), name);
  }
  const out = ctx.XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}

// Full pipeline: sheets -> xlsx bytes -> parse -> clean -> computeModel
function modelFromSheets(ctx, sheets) {
  const wb = ctx.XLSX.read(sheetsToBytes(ctx, sheets), { type: 'array', cellDates: false });
  return ctx.__app.computeModel(ctx.__app.cleanData(wb));
}

// Same but starting from an in-memory workbook object (e.g. the sample builder)
function modelFromWorkbook(ctx, wb) {
  const out = ctx.XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const bytes = out instanceof Uint8Array ? out : new Uint8Array(out);
  const wb2 = ctx.XLSX.read(bytes, { type: 'array', cellDates: false });
  return ctx.__app.computeModel(ctx.__app.cleanData(wb2));
}

module.exports = { makeCtx, sheetsToBytes, modelFromSheets, modelFromWorkbook };
