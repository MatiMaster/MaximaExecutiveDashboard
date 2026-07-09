/* ============================================================
 * Section 7 — product classification. Growth = period sales vs. same period
 * last year; buckets are doing-great (>= +15%), OK (strictly between), and
 * slow/declining (<= -15%). No prior-year sales => 0% => OK. Products without
 * a segment or with no activity are excluded. `adjusted` swaps in the
 * backorder-adjusted total (YTD + backorder) — the section's toggle.
 * classifyProducts is a pure function, so we test it directly.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx } = require('./helper');

const { classifyProducts } = makeCtx('es').__app;

// item, seg, ytd, ly, units, bo  ->  the fields classifyProducts reads
const P = (item, seg, ytd, ly, bo = 0) => ({ item, name: item, seg, ytd, ly, units: 1, bo });

test('buckets by ±15% growth vs last year (boundaries inclusive on the outside)', () => {
  const c = classifyProducts([
    P('GROW', 'S', 200, 100),   // +100% -> great
    P('EDGEG', 'S', 115, 100),  // +15%  -> great (>=)
    P('FLATP', 'S', 110, 100),  // +10%  -> OK
    P('FLATN', 'S', 90, 100),   // -10%  -> OK
    P('EDGES', 'S', 85, 100),   // -15%  -> slow (<=)
    P('DROP', 'S', 40, 100)     // -60%  -> slow
  ], false);
  assert.equal(c.analyzed, 6);
  assert.deepEqual(c.good.map(r => r.item), ['GROW', 'EDGEG']);
  assert.deepEqual(c.ok.map(r => r.item).sort(), ['FLATN', 'FLATP']);
  assert.deepEqual(c.slow.map(r => r.item).sort(), ['DROP', 'EDGES']);
});

test('no prior-year sales counts as 0% (OK), not as growth', () => {
  const c = classifyProducts([P('NEW', 'S', 5000, 0)], false);
  assert.equal(c.ok.length, 1);
  assert.equal(c.good.length, 0);
  assert.equal(c.ok[0].pct, 0);
});

test('products without a segment, or with no activity, are excluded', () => {
  const c = classifyProducts([
    P('NOSEG', '', 500, 100),   // excluded: no segment
    P('DEAD', 'S', 0, 0),       // excluded: no sales and no LY
    P('LIVE', 'S', 500, 100)    // kept
  ], false);
  assert.equal(c.analyzed, 1);
  assert.equal(c.good[0].item, 'LIVE');
});

test('rank is assigned by sales descending across the analyzed set', () => {
  const c = classifyProducts([
    P('SMALL', 'S', 100, 50),
    P('BIG', 'S', 900, 50),
    P('MID', 'S', 500, 50)
  ], false);
  const ranks = Object.fromEntries([...c.good, ...c.ok, ...c.slow].map(r => [r.item, r.rank]));
  assert.equal(ranks.BIG, 1);
  assert.equal(ranks.MID, 2);
  assert.equal(ranks.SMALL, 3);
});

test('backorder-adjusted view can move a product into a different bucket', () => {
  const rows = [P('X', 'S', 90, 100, /*bo*/ 60)]; // actual -10% (OK); adjusted (90+60)/100 = +50% (great)
  assert.equal(classifyProducts(rows, false).ok.length, 1);
  const adj = classifyProducts(rows, true);
  assert.equal(adj.good.length, 1);
  assert.equal(adj.good[0].sales, 150);
});

test('every analyzed product lands in exactly one bucket', () => {
  const c = classifyProducts([
    P('A', 'S', 200, 100), P('B', 'S', 100, 100), P('C', 'S', 50, 100), P('D', 'S', 300, 0)
  ], false);
  assert.equal(c.good.length + c.ok.length + c.slow.length, c.analyzed);
});
