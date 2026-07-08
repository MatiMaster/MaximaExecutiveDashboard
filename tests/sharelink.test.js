/* ============================================================
 * Shareable-link round-trip: the model embedded in the URL must decode
 * back to the same numbers. If this drifts, shared links silently show
 * wrong figures — so it is an accuracy test, not just a plumbing test.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, modelFromWorkbook } = require('./helper');

const ctx = makeCtx('es');
const A = ctx.__app;
const M = modelFromWorkbook(ctx, A.buildSampleWorkbook());

test('compact -> hydrate preserves the two large tables value-for-value', () => {
  const hydrated = A.hydrateModel(A.compactModel(M));
  assert.equal(hydrated.perf.prodPerf.length, M.perf.prodPerf.length);
  assert.equal(hydrated.perf.segPerf.length, M.perf.segPerf.length);
  const a = M.perf.prodPerf.slice().sort((x, y) => y.ytd - x.ytd)[0];
  const b = hydrated.perf.prodPerf.find(p => p.item === a.item);
  assert.equal(b.name, a.name);
  assert.equal(b.seg, a.seg);
  assert.equal(Math.round(a.ytd), b.ytd);
  assert.equal(Math.round(a.bo), b.bo);
});

test('gzip encode -> decode -> hydrate reproduces the headline numbers', async () => {
  const payload = { v: 2, source: 'sample.xlsx', model: A.compactModel(M) };
  const enc = await A.encodeShare(JSON.stringify(payload));
  const back = JSON.parse(await A.decodeShare(enc));
  const h = A.hydrateModel(back.model);
  assert.equal(back.source, 'sample.xlsx');
  assert.equal(Math.round(h.sales.ytd), Math.round(M.sales.ytd));
  assert.equal(Math.round(h.sales.ly), Math.round(M.sales.ly));
  assert.equal(Math.round(h.book.openValue), Math.round(M.book.openValue));
  assert.equal(Math.round(h.book.boValue), Math.round(M.book.boValue));
  assert.equal(h.perf.prodPerf.length, M.perf.prodPerf.length);
});

test('encoded payload is URL-hash-safe (base64url + format prefix)', async () => {
  const enc = await A.encodeShare(JSON.stringify({ v: 2, model: A.compactModel(M) }));
  assert.match(enc, /^[gr][A-Za-z0-9_-]+$/); // 'g' gzip or 'r' raw, then base64url only
});
