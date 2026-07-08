/* ============================================================
 * Data-cleaning rules that gate accuracy: source exclusions, the
 * Amount (not foreign-currency) column, warehouse normalization,
 * segment back-fill, date parsing, and the same-period-last-year shift.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, modelFromSheets } = require('./helper');
const { SHEETS, EXPECTED } = require('./fixture');

const ctx = makeCtx('es');

test('cleaning: Oscar Riquelme / CMM02359 / Subtotal rows are excluded', () => {
  const wb = ctx.XLSX.read(require('./helper').sheetsToBytes(ctx, SHEETS), { type: 'array' });
  const data = ctx.__app.cleanData(wb);
  assert.equal(data.sales.length, EXPECTED.cleaning.salesRows);
  assert.equal(data.sales.filter(r => /oscar riquelme/i.test(r.rep)).length, 0);
  assert.equal(data.sales.filter(r => r.id.includes('CMM02359') || r.c.includes('CMM02359')).length, 0);
  assert.equal(data.sales.filter(r => r.item.toLowerCase() === 'subtotal').length, 0);
});

test('cleaning: uses the Amount column, never Amount (Foreign Currency)', () => {
  // every fixture row has foreign=999; YTD would be huge if the wrong column won
  const M = modelFromSheets(ctx, SHEETS);
  assert.equal(Math.round(M.sales.ytd), EXPECTED.sales.ytd);
  assert.ok(M.sales.ytd < 5000, 'foreign-currency decoy did not leak in');
});

test('cleaning: VTNWH-MX-EU normalizes to VTNWH-MX', () => {
  const M = modelFromSheets(ctx, SHEETS);
  const whNames = [...M.book.readyWh, ...M.book.boWh].map(([k]) => k);
  assert.ok(whNames.includes('VTNWH-MX'));
  assert.ok(!whNames.includes('VTNWH-MX-EU'));
});

test('cleaning: product segment is back-filled onto PendingFullFill from Sales', () => {
  // PFF sheet has no segment column; segments must come from the item->segment map
  const M = modelFromSheets(ctx, SHEETS);
  assert.equal(Math.round(M.perf.segPerf.find(s => s.seg === '4T Engine Oils').bo), EXPECTED.segPerf['4T Engine Oils'].bo);
  assert.equal(Math.round(M.perf.segPerf.find(s => s.seg === 'Aerosols').bo), EXPECTED.segPerf['Aerosols'].bo);
});

test('date parsing: ISO strings, Excel serials, and Date objects agree', () => {
  const { toISO } = ctx.__app;
  assert.equal(toISO('2026-02-10'), '2026-02-10');
  assert.equal(toISO('2026-02-10T00:00:00'), '2026-02-10');
  // Excel serial round-trip (1900 date system): serial 25569 == 1970-01-01
  const serial = Math.round(Date.UTC(2026, 1, 10) / 86400000) + 25569;
  assert.equal(toISO(serial), '2026-02-10');
  assert.equal(toISO(''), '');
  assert.equal(toISO(null), '');
});

test('same-period-last-year shift subtracts exactly one year', () => {
  const { lyISO } = ctx.__app;
  assert.equal(lyISO('2026-07-01'), '2025-07-01');
  assert.equal(lyISO('2026-01-01'), '2025-01-01');
});
