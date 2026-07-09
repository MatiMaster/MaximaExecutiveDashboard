/* ============================================================
 * Filters — narrowing the cleaned data must flow through to EVERY
 * number. These assert the sales KPIs, the order book, and the funnel
 * all respond to each dimension, and that "no filters" reproduces the
 * full-data model exactly. Expected values are derived by hand from the
 * fixture rows (see tests/fixture.js).
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, filteredModelFromSheets } = require('./helper');
const { SHEETS, EXPECTED } = require('./fixture');

const ctx = makeCtx('es');
const F = over => Object.assign(ctx.__app.emptyFilters(), over);
const model = over => filteredModelFromSheets(ctx, SHEETS, F(over));
const ytdOf = over => Math.round(model(over).sales.ytd);
const arr = x => [...x];  // re-home VM-realm arrays so deepStrictEqual compares by value

test('filters: client category derives from the sales rep', () => {
  const cc = ctx.__app.clientCat;
  assert.equal(cc('Andres Orozco'), 'International');
  assert.equal(cc('Cristian Calderon'), 'International');
  assert.equal(cc('Rep One'), 'Domestic');        // unknown rep -> Domestic
  assert.equal(cc('  ANDRES  orozco '), 'International'); // accent/space/case-insensitive
});

test('filters: empty filters reproduce the full-data model exactly', () => {
  const m = model({});
  assert.equal(Math.round(m.sales.ytd), EXPECTED.sales.ytd);       // 1050
  assert.equal(Math.round(m.sales.ly), EXPECTED.sales.ly);         // 80
  assert.equal(Math.round(m.book.openValue), EXPECTED.book.openValue); // 1300
  assert.equal(m.sales.activeCustomers, EXPECTED.sales.activeCustomers);
  assert.deepEqual(arr(m.sales.monthly.cur), EXPECTED.sales.monthlyCur);
});

test('filters: Sales Rep narrows sales KPIs (YTD, LY, customers)', () => {
  const m = model({ rep: ['Rep One'] });   // INV1+INV2 (2026), INV5 (LY)
  assert.equal(Math.round(m.sales.ytd), 300);   // 100 + 200
  assert.equal(Math.round(m.sales.ly), 80);     // INV5
  assert.equal(m.sales.activeCustomers, 1);     // Cust A
  assert.equal(m.sales.countries, 1);           // Chile
});

test('filters: Customer narrows to that customer only', () => {
  assert.equal(ytdOf({ customer: ['Cust B'] }), 350);  // INV3 300 + INV4 50
});

test('filters: Country / Customer Type / Customer Category each subset sales', () => {
  assert.equal(ytdOf({ country: ['Chile'] }), 300);      // INV1+INV2
  assert.equal(ytdOf({ ctype: ['EX-LATAM'] }), 350);     // Cust B: INV3+INV4
  assert.equal(ytdOf({ custCat: ['CatX'] }), 700);       // INV1+INV2+INV10
});

test('filters: Client Category=Domestic keeps all (fixture reps are all domestic)', () => {
  assert.equal(ytdOf({ clientCat: ['Domestic'] }), EXPECTED.sales.ytd);   // 1050
  assert.equal(ytdOf({ clientCat: ['International'] }), 0);               // none
});

test('filters: Product Segment flows into BOTH sales and the order book', () => {
  const m = model({ seg: ['2T Engine Oils'] });   // sales INV2+INV10; PFF only SO2
  assert.equal(Math.round(m.sales.ytd), 600);        // 200 + 400
  assert.equal(Math.round(m.book.openValue), 500);   // SO2 aggregate only
  assert.equal(Math.round(m.book.boValue), 0);       // SO2 has no backorder
  assert.equal(m.book.funnel.open.qty, 1);           // one SO left in the book
});

test('filters: Location (warehouse) subsets sales and the order book together', () => {
  const m = model({ loc: ['EU Warehouse'] });   // sales INV3+INV4; PFF only SO2
  assert.equal(Math.round(m.sales.ytd), 350);
  assert.equal(Math.round(m.book.openValue), 500);
});

test('filters: date range redefines the period and shifts LY with it', () => {
  const m = model({ from: '2026-03-01', to: '2026-06-30' });
  // in-window 2026: INV2 (Mar) 200 + INV10 (May) 400 + INV3 (Jun) 300 = 900
  // out: INV1 (Feb, before From), INV4 (Jul, after To)
  assert.equal(Math.round(m.sales.ytd), 900);
  // LY window [2025-03-01, 2025-06-30] contains no 2025 rows (INV5 is Feb)
  assert.equal(Math.round(m.sales.ly), 0);
  assert.equal(m.asOf, '2026-06-30');          // period end drives the as-of label
  assert.equal(m.yearStart, '2026-03-01');     // period start
});

test('filters: date range also clips the open book by SO date', () => {
  // all fixture PFF rows are dated May 2026 -> inside a wide window, outside a narrow one
  assert.equal(Math.round(model({ from: '2026-05-01', to: '2026-05-31' }).book.openValue), 1300);
  assert.equal(Math.round(model({ from: '2026-01-01', to: '2026-04-30' }).book.openValue), 0);
});

test('filters: dimensions compose (Rep One AND 4T segment => INV1 only)', () => {
  const m = model({ rep: ['Rep One'], seg: ['4T Engine Oils'] });
  assert.equal(Math.round(m.sales.ytd), 100);   // INV1
  assert.equal(Math.round(m.sales.ly), 80);     // INV5 (4T, Rep One, LY)
});

test('filters: Document Number does an exact match on sales', () => {
  assert.equal(ytdOf({ doc: 'INV1' }), 100);
});

test('filterOptions: distinct sorted menus come from the cleaned data', () => {
  const wb = ctx.XLSX.read(require('./helper').sheetsToBytes(ctx, SHEETS), { type: 'array' });
  const opts = ctx.__app.filterOptions(ctx.__app.cleanData(wb));
  assert.deepEqual(arr(opts.country), ['Chile', 'Germany', 'Japan']);
  assert.deepEqual(arr(opts.seg), ['2T Engine Oils', '4T Engine Oils', 'Aerosols']);
  assert.deepEqual(arr(opts.ctype), ['EX-ASIA', 'EX-LATAM']);
  assert.ok(opts.rep.includes('Rep One') && opts.rep.includes('Rep Two'));
  assert.ok(!opts.rep.some(r => /oscar/i.test(r)));  // excluded rep never offered
});

test('filterData: filtering to a SO document empties sales but keeps that SO in the book', () => {
  const wb = ctx.XLSX.read(require('./helper').sheetsToBytes(ctx, SHEETS), { type: 'array' });
  const fd = ctx.__app.filterData(ctx.__app.cleanData(wb), F({ doc: 'SO1' }));
  assert.equal(fd.sales.length, 0);                       // no sales invoice is "SO1"
  assert.ok(fd.pff.length > 0 && fd.pff.every(r => r.doc === 'SO1'));
});
