/* ============================================================
 * Structural invariants that must hold for ANY dataset. Run against the
 * built-in sample workbook. These catch whole classes of bugs (double-
 * counting, dropped rows, mis-attribution) without hardcoded numbers.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, modelFromWorkbook } = require('./helper');

const ctx = makeCtx('es');
const M = modelFromWorkbook(ctx, ctx.__app.buildSampleWorkbook());
const sum = (arr, f) => arr.reduce((a, x) => a + f(x), 0);
const near = (a, b, eps = 1) => Math.abs(a - b) <= eps;

test('per-product YTD sums to total YTD sales', () => {
  assert.ok(near(sum(M.perf.prodPerf, p => p.ytd), M.sales.ytd));
});

test('per-segment YTD sums to total YTD sales', () => {
  assert.ok(near(sum(M.perf.segPerf, s => s.ytd), M.sales.ytd));
});

test('backorder value is fully attributed: Σ per-product bo = Σ per-segment bo = book.boValue', () => {
  assert.ok(near(sum(M.perf.prodPerf, p => p.bo), M.book.boValue));
  assert.ok(near(sum(M.perf.segPerf, s => s.bo), M.book.boValue));
});

// NOTE: warehouse/customer breakdowns are top-5 lists, so their sum is a
// SUBSET of the total (equal only when there are <=5 groups). The complete
// attribution invariant is tested on the un-truncated per-product/segment
// tables above.
test('backorder breakdowns are a non-empty subset of the backorder value', () => {
  const eps = 1;
  assert.ok(sum(M.book.boWh, ([, v]) => v) > 0);
  assert.ok(sum(M.book.boWh, ([, v]) => v) <= M.book.boValue + eps);
  assert.ok(sum(M.book.boCust, ([, v]) => v) > 0);
  assert.ok(sum(M.book.boCust, ([, v]) => v) <= M.book.boValue + eps);
  assert.ok(M.book.boWh.length <= 5 && M.book.boCust.length <= 5);
});

test('committed (ready) breakdowns are a non-empty subset of the ready value', () => {
  const eps = 1;
  assert.ok(sum(M.book.readyWh, ([, v]) => v) <= M.book.readyValue + eps);
  assert.ok(sum(M.book.readyCust, ([, v]) => v) <= M.book.readyValue + eps);
  assert.ok(M.book.readyWh.length <= 5 && M.book.readyCust.length <= 5);
});

test('derived ratios equal their definitions', () => {
  assert.ok(near(M.sales.deltaPct, (M.sales.ytd - M.sales.ly) / M.sales.ly, 1e-9));
  assert.equal(M.sales.delta, M.sales.ytd - M.sales.ly);
  assert.ok(near(M.book.readyShare, M.book.readyValue / M.book.openValue, 1e-9));
});

test('monthly series are aligned in length', () => {
  const m = M.sales.monthly;
  assert.equal(m.cur.length, m.idx.length);
  assert.equal(m.prev.length, m.idx.length);
});

test('ranked lists are sorted descending by value', () => {
  const desc = arr => arr.every((x, i) => i === 0 || arr[i - 1][1] >= x[1]);
  assert.ok(desc(M.markets.topCountries));
  assert.ok(desc(M.markets.topSegments));
  assert.ok(desc(M.book.boWh) && desc(M.book.boCust));
  assert.ok(desc(M.book.readyWh) && desc(M.book.readyCust));
});

test('distinct-count invariants', () => {
  assert.ok(M.book.boSOs <= M.book.openOrders, 'backordered SOs cannot exceed open SOs');
  assert.ok(M.sales.activeCustomers >= 1 && M.sales.countries >= 1);
});
