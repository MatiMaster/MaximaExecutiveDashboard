/* ============================================================
 * Number/percent formatters — these decide what the executive actually
 * READS, so they're part of value accuracy. Tested in both locales
 * (Spanish uses "." thousands / "," decimals; English the reverse).
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx } = require('./helper');

const es = makeCtx('es').__app;
const en = makeCtx('en').__app;

test('full(): whole-dollar with locale grouping', () => {
  assert.equal(es.full(1050), '$1.050');
  assert.equal(en.full(1050), '$1,050');
  assert.equal(es.full(1234567), '$1.234.567');
  assert.equal(en.full(1234567), '$1,234,567');
  assert.equal(es.full(0), '$0');
  assert.equal(es.full(499.6), '$500');      // rounds
});

test('full(): negatives render as -$N (not $-N)', () => {
  assert.equal(es.full(-450), '-$450');
  assert.equal(en.full(-450), '-$450');
});

test('short(): millions / thousands / plain, with es comma decimals', () => {
  assert.equal(es.short(2533852), '$2,53M');
  assert.equal(en.short(2533852), '$2.53M');
  assert.equal(es.short(450000), '$450K');
  assert.equal(es.short(450), '$450');       // below 1k falls back to full()
  assert.equal(es.short(-2533852), '-$2,53M');
});

test('num(): rounded integer with grouping', () => {
  assert.equal(es.num(58), '58');
  assert.equal(es.num(1050), '1.050');
  assert.equal(en.num(1050), '1,050');
});

test('pct1(): one decimal, locale separator, signed by input', () => {
  assert.equal(es.pct1(0.5714285), '57,1%');
  assert.equal(en.pct1(0.5714285), '57.1%');
  assert.equal(es.pct1(-0.028), '-2,8%');
  assert.equal(es.pct1(0), '0,0%');
});

test('pct0(): rounded whole percent', () => {
  assert.equal(es.pct0(0.6538), '65%');
  assert.equal(es.pct0(1), '100%');
});
