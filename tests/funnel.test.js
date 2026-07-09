/* ============================================================
 * Fulfillment funnel (Section 3): the open book split into
 *   Open SO  ->  Waiting IF  /  IF In Progress  ->  BO
 * The split hinges on joining PendingFullFill to the IF sheet by the
 * IF's "Created" column (its source SO). These asserts pin every card.
 * ============================================================ */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { makeCtx, modelFromSheets } = require('./helper');
const { SHEETS, EXPECTED } = require('./fixture');

const ctx = makeCtx('es');
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-9, `${a} !~= ${b}`);

test('funnel: Open SO = the whole open book', () => {
  const f = modelFromSheets(ctx, SHEETS).book.funnel;
  const e = EXPECTED.book.funnel.open;
  assert.equal(f.open.qty, e.qty);
  assert.equal(Math.round(f.open.amount), e.amount);
  near(f.open.avgDays, e.avgDays);
});

test('funnel: Waiting IF = pending SOs with no Item Fulfillment yet', () => {
  const f = modelFromSheets(ctx, SHEETS).book.funnel;
  const e = EXPECTED.book.funnel.waiting;
  assert.equal(f.waiting.qty, e.qty);
  assert.equal(Math.round(f.waiting.amount), e.amount);
  near(f.waiting.avgDays, e.avgDays);
});

test('funnel: IF In Progress = pending SOs an IF was created from', () => {
  const f = modelFromSheets(ctx, SHEETS).book.funnel;
  const e = EXPECTED.book.funnel.inProg;
  assert.equal(f.inProg.ifQty, e.ifQty);   // distinct IF documents
  assert.equal(f.inProg.soQty, e.soQty);   // distinct source SOs
  assert.equal(Math.round(f.inProg.amount), e.amount);
  near(f.inProg.avgDays, e.avgDays);        // avg over IF dates, not SO dates
});

test('funnel: BO card matches the backorder totals + % of open value', () => {
  const f = modelFromSheets(ctx, SHEETS).book.funnel;
  const e = EXPECTED.book.funnel.bo;
  assert.equal(f.bo.units, e.units);
  assert.equal(Math.round(f.bo.amount), e.amount);
  assert.equal(f.bo.sos, e.sos);
  near(f.bo.avgDays, e.avgDays);
  near(f.bo.pctOfOpen, e.pctOfOpen);
});

test('funnel: Open SO amount = Waiting IF + IF In Progress (a clean partition)', () => {
  const f = modelFromSheets(ctx, SHEETS).book.funnel;
  near(f.open.amount, f.waiting.amount + f.inProg.amount);
  assert.equal(f.open.qty, f.waiting.qty + f.inProg.soQty);
});
