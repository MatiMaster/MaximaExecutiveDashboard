/* ============================================================
 * Every value the executive dashboard DISPLAYS, checked against the
 * hand-computed fixture. Grouped by dashboard section.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, modelFromSheets } = require('./helper');
const { SHEETS, EXPECTED } = require('./fixture');

const ctx = makeCtx('es');
// computeModel returns objects from the VM realm; normalize into this realm so
// deepStrictEqual compares by structure, not prototype identity.
const M = JSON.parse(JSON.stringify(modelFromSheets(ctx, SHEETS)));
const round = n => Math.round(n * 100) / 100;
const item = sku => M.perf.prodPerf.find(p => p.item === sku);
const seg = name => M.perf.segPerf.find(s => s.seg === name);

test('as-of date + reporting year derive from the latest Sales date', () => {
  assert.equal(M.asOf, EXPECTED.asOf);
  assert.equal(M.year, EXPECTED.year);
  assert.equal(M.prev, EXPECTED.prev);
});

test('Section 1 — sales KPIs (YTD, same-period LY, delta, customers, countries)', () => {
  const e = EXPECTED.sales;
  assert.equal(round(M.sales.ytd), e.ytd);
  assert.equal(round(M.sales.ly), e.ly);
  assert.equal(round(M.sales.delta), e.delta);
  assert.equal(round(M.sales.deltaPct), round(e.deltaPct));
  assert.equal(M.sales.activeCustomers, e.activeCustomers);
  assert.equal(M.sales.countries, e.countries);
});

test('Section 1 — monthly bars: complete months only; partial month excluded but kept in YTD', () => {
  const e = EXPECTED.sales, m = M.sales.monthly;
  assert.deepEqual(m.idx, e.monthlyIdx);
  assert.deepEqual(m.cur.map(round), e.monthlyCur);
  assert.deepEqual(m.prev.map(round), e.monthlyPrev);
  assert.equal(m.partialMonth, e.partialMonth);
  assert.equal(round(m.cur.reduce((a, b) => a + b, 0)), e.monthlyCurSum);
  assert.notEqual(e.monthlyCurSum, e.ytd); // the July line is in YTD, not in the bars
});

test('Section 2 — markets: top countries & segments by YTD amount, plus shares', () => {
  const e = EXPECTED.markets;
  assert.deepEqual(M.markets.topCountries.map(([k, v]) => [k, round(v)]), e.topCountries);
  assert.deepEqual(M.markets.topSegments.map(([k, v]) => [k, round(v)]), e.topSegments);
  assert.equal(round(M.markets.topCountriesShare), round(e.topCountriesShare));
  assert.equal(round(M.markets.topSegShare), round(e.topSegShare));
});

test('Section 3 — open order book (value, counts, composition, status, units)', () => {
  const e = EXPECTED.book, b = M.book;
  assert.equal(round(b.openValue), e.openValue);
  assert.equal(b.openOrders, e.openOrders);
  assert.equal(b.bookCustomers, e.bookCustomers);
  assert.equal(round(b.readyValue), e.readyValue);
  assert.equal(round(b.boValue), e.boValue);
  assert.equal(round(b.orderedUnits), e.orderedUnits);
  assert.equal(round(b.readyShare), round(e.readyShare));
  assert.deepEqual(b.status, e.status);
});

test('Section 4 — backorders (value, units, SOs, by warehouse, by customer)', () => {
  const e = EXPECTED.book, b = M.book;
  assert.equal(round(b.boValue), e.boValue);
  assert.equal(round(b.boUnits), e.boUnits);
  assert.equal(b.boSOs, e.boSOs);
  assert.deepEqual(b.boWh.map(([k, v]) => [k, round(v)]), e.boWh);
  assert.deepEqual(b.boCust.map(([k, v]) => [k, round(v)]), e.boCust);
});

test('Section 5 — awaiting fulfillment (committed value/units, share, by warehouse/customer)', () => {
  const e = EXPECTED.book, b = M.book;
  assert.equal(round(b.readyValue), e.readyValue);
  assert.equal(round(b.readyUnits), e.readyUnits);
  assert.equal(round(b.readyShare), round(e.readyShare));
  assert.deepEqual(b.readyWh.map(([k, v]) => [k, round(v)]), e.readyWh);
  assert.deepEqual(b.readyCust.map(([k, v]) => [k, round(v)]), e.readyCust);
});

test('Section 6 — per-product performance (YTD/units/LY/FY/backorder, name & segment)', () => {
  for (const [sku, e] of Object.entries(EXPECTED.prodPerf)) {
    const p = item(sku);
    assert.ok(p, `product ${sku} present`);
    assert.equal(round(p.ytd), e.ytd, `${sku} ytd`);
    assert.equal(round(p.units), e.units, `${sku} units`);
    assert.equal(round(p.ly), e.ly, `${sku} ly`);
    assert.equal(round(p.fy), e.fy, `${sku} fy`);
    assert.equal(round(p.bo), e.bo, `${sku} bo`);
    assert.equal(round(p.boUnits), e.boUnits, `${sku} boUnits`);
    assert.equal(p.seg, e.seg, `${sku} segment (back-filled)`);
    assert.equal(p.name, e.name, `${sku} name`);
  }
});

test('Section 6 — the backorder-adjusted total is YTD + backorder, per product', () => {
  for (const [sku, e] of Object.entries(EXPECTED.prodPerf)) {
    const p = item(sku);
    assert.equal(round(p.ytd + p.bo), round(e.ytd + e.bo), `${sku} adjusted`);
  }
});

test('Section 6 — per-segment performance (YTD/units/LY/FY/backorder)', () => {
  for (const [name, e] of Object.entries(EXPECTED.segPerf)) {
    const s = seg(name);
    assert.ok(s, `segment ${name} present`);
    assert.equal(round(s.ytd), e.ytd, `${name} ytd`);
    assert.equal(round(s.units), e.units, `${name} units`);
    assert.equal(round(s.ly), e.ly, `${name} ly`);
    assert.equal(round(s.fy), e.fy, `${name} fy`);
    assert.equal(round(s.bo), e.bo, `${name} bo`);
  }
});

test('per-product deltas: displayed Δ = value - baseline (actual and adjusted)', () => {
  const p = item('SKU1'); // ytd 400, ly 80, fy 200, bo 250
  assert.equal(round(p.ytd - p.ly), 320);        // Δ vs same-period LY
  assert.equal(round((p.ytd - p.ly) / p.ly), 4); // % Δ = 320/80
  assert.equal(round(p.ytd - p.fy), 200);        // Δ vs full prior year
  assert.equal(round(p.ytd + p.bo - p.ly), 570); // adjusted Δ vs LY = 650-80
});
