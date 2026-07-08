/* ============================================================
   MRO International — Executive Dashboard (dynamic)
   Upload an .xlsx (Sales / Target / PendingFullFill / IF), parse
   it in the browser, and render the executive view. Every figure
   is computed deterministically from the file — no stored numbers.
   ============================================================ */
'use strict';

/* ------------------------------------------------------------------ *
 * 0. Global state
 * ------------------------------------------------------------------ */
let LANG = localStorage.getItem('mro.lang') || 'es';   // default: Spanish
let THEME = localStorage.getItem('mro.theme') || 'light'; // default: clear
let MODEL = null;      // computed metrics
let SOURCE = '';       // uploaded file name
let ITEM_LIMIT = 25;   // top-N for the product performance table
let BO_SEG = false;    // segment table: false = actual · true = backorder-adjusted
let BO_PROD = false;   // product table: false = actual · true = backorder-adjusted

const $ = id => document.getElementById(id);
const S = () => I18N.STR[LANG];
const fmt = (tpl, vars) => String(tpl).replace(/\{(\w+)\}/g, (_, k) => (vars && k in vars) ? vars[k] : '');

/* ------------------------------------------------------------------ *
 * 1. Formatting (locale-aware)
 * ------------------------------------------------------------------ */
const locale = () => LANG === 'es' ? 'es-CL' : 'en-US';
const full = n => { n = Math.round(Number(n) || 0); return (n < 0 ? '-' : '') + '$' + Math.abs(n).toLocaleString(locale()); };
const signed = n => (Number(n) > 0 ? '+' : '') + full(n);
const num = n => Math.round(Number(n) || 0).toLocaleString(locale());
function short(n) {
  n = Number(n) || 0; const a = Math.abs(n);
  if (a >= 1e6) { let s = (a / 1e6).toFixed(2); if (LANG === 'es') s = s.replace('.', ','); return (n < 0 ? '-' : '') + '$' + s + 'M'; }
  if (a >= 1e3) { return (n < 0 ? '-' : '') + '$' + Math.round(a / 1e3) + 'K'; }
  return full(n);
}
function pct1(x) { if (!isFinite(x)) x = 0; let s = (x * 100).toFixed(1); if (LANG === 'es') s = s.replace('.', ','); return s + '%'; }
function pct0(x) { if (!isFinite(x)) x = 0; return Math.round(x * 100) + '%'; }

// short "day month" e.g. "1 ene" / "Jan 1"
function shortDM(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return LANG === 'es' ? `${d} ${I18N.MONTHS.es[m - 1].toLowerCase()}` : `${I18N.MONTHS.en[m - 1]} ${d}`;
}
// long date e.g. "2 de julio de 2026" / "July 2, 2026"
function longDate(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return LANG === 'es' ? `${d} de ${I18N.MONTHS_LONG.es[m - 1]} de ${y}` : `${I18N.MONTHS_LONG.en[m - 1]} ${d}, ${y}`;
}

/* ------------------------------------------------------------------ *
 * 2. Parsing helpers
 * ------------------------------------------------------------------ */
const normHeader = h => String(h == null ? '' : h).replace(/\s+/g, ' ').trim().toLowerCase();
const repKey = v => String(v || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().replace(/\s+/g, ' ').toLowerCase();
const locNorm = v => { const s = String(v == null ? '' : v).trim(); return (s === 'VTNWH-MX-EU' || s === 'VTNWH-MX') ? 'VTNWH-MX' : s; };
const N = v => { if (typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; };
const T = v => (v == null ? '' : String(v)).trim();

// Excel serial (1900 system) -> ISO 'YYYY-MM-DD', deterministic via UTC.
function serialToISO(serial) {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
}
function toISO(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return serialToISO(v);
  if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  return '';
}
// shift ISO date back one year (for same-period-last-year windows)
function lyISO(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return (Number(y) - 1) + '-' + m + '-' + d;
}

// Find a worksheet by any of several candidate names (case-insensitive).
function findSheet(wb, names) {
  const want = names.map(n => n.toLowerCase());
  for (const nm of wb.SheetNames) if (want.includes(nm.toLowerCase())) return wb.Sheets[nm];
  return null;
}

// Turn a sheet into {rows} where each row is an object keyed by our field
// names, using a per-field list of acceptable header labels.
function readSheet(ws, fieldMap) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  if (!aoa.length) return [];
  const headers = aoa[0].map(normHeader);
  const idx = {}; // field -> column index
  for (const [field, cands] of Object.entries(fieldMap)) {
    for (const cand of cands) {
      const c = headers.indexOf(normHeader(cand));
      if (c !== -1) { idx[field] = c; break; }
    }
  }
  const out = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.every(c => c == null || c === '')) continue;
    const o = {};
    for (const field of Object.keys(fieldMap)) o[field] = field in idx ? row[idx[field]] : null;
    out.push(o);
  }
  return out;
}

const FIELDS = {
  sales: {
    id: ['ID'], c: ['Customer'], rep: ['Sales Rep'], tt: ['Transaction Type'], date: ['Date'],
    doc: ['Document Number'], item: ['Item'], memo: ['Memo'],
    sub: ['Subsidiary: Name', 'Subsidiary'], cat: ['Customer Category: Name', 'Customer Category'],
    q: ['Quantity'], up: ['Unit Price'], amt: ['Amount'],
    ctype: ['Customer/Project: Customer Type', 'Customer Type'],
    loc: ['Inventory Location: Name', 'Inventory Location', 'Location'],
    country: ['Address: Billing Address Country Name', 'Billing Address Country Name', 'Country'],
    seg: ['Product Segment', 'Segment'], sod: ['SO Date']
  },
  pff: {
    itype: ['Item Type'], item: ['Item'], desc: ['Description (Sales)', 'Description'], date: ['Date'],
    tt: ['Transaction Type'], rep: ['Primary Sales Rep', 'Sales Rep'], doc: ['Document Number'], c: ['Customer'],
    ord: ['Ordered'], ful: ['Fulfilled'], com: ['Committed'], bo: ['Back Ordered', 'Backordered'],
    sub: ['Subsidiary: Name', 'Subsidiary'], agg: ['Aggregate Amount'], up: ['Unit Price'],
    loc: ['Location', 'Inventory Location: Name'], vstatus: ['Validated Status', 'Status'], seg: ['Product Segment', 'Segment']
  },
  target: { id: ['ID'], amt: ['Amount'], country: ['Country'], ctype: ['Customer Type'], type: ['Type'] },
  ifs: {
    iid: ['Internal ID'], sub: ['Subsidiary'], date: ['Date'], type: ['Type'], doc: ['Document Number'],
    name: ['Name'], status: ['Status'], amt: ['Amount'], cat: ['Category'], rep: ['Sales Rep'],
    loc: ['Location'], country: ['Shipping Country'], created: ['Created']
  }
};

/* ------------------------------------------------------------------ *
 * 3. Cleaning + model
 * ------------------------------------------------------------------ */
function cleanData(wb) {
  const wsSales = findSheet(wb, ['Sales']);
  const wsPff = findSheet(wb, ['PendingFullFill', 'PendingFulFill', 'Pending Fulfillment', 'Pending']);
  const wsTarget = findSheet(wb, ['Target', '26_Target']);
  const wsIf = findSheet(wb, ['IF']);
  if (!wsSales) throw { key: 'errNoSheet', sheet: 'Sales' };
  if (!wsPff) throw { key: 'errNoSheet', sheet: 'PendingFullFill' };

  // ---- Sales: apply the source exclusion rules ----
  let sales = readSheet(wsSales, FIELDS.sales).map(r => ({
    id: T(r.id), c: T(r.c), rep: T(r.rep), tt: T(r.tt), d: toISO(r.date),
    doc: T(r.doc), item: T(r.item), memo: T(r.memo), sub: T(r.sub), cat: T(r.cat),
    q: N(r.q), up: N(r.up), amt: N(r.amt), ctype: T(r.ctype), loc: locNorm(r.loc),
    country: T(r.country), seg: T(r.seg), sod: toISO(r.sod)
  })).filter(r => {
    if (repKey(r.rep) === 'oscar riquelme') return false;                 // excluded rep
    if (r.id.includes('CMM02359') || r.c.includes('CMM02359')) return false; // excluded account
    if (r.item.toLowerCase() === 'subtotal') return false;                // subtotal lines
    return r.d || r.amt || r.c || r.id;                                   // drop empties
  });
  if (!sales.length) throw { key: 'errNoSales' };

  // ---- Pending Fulfillment ----
  let pff = readSheet(wsPff, FIELDS.pff).map(r => ({
    itype: T(r.itype), item: T(r.item), desc: T(r.desc), d: toISO(r.date), tt: T(r.tt),
    rep: T(r.rep), doc: T(r.doc), c: T(r.c), ord: N(r.ord), ful: N(r.ful), com: N(r.com),
    bo: N(r.bo), sub: T(r.sub), agg: N(r.agg), up: N(r.up), loc: locNorm(r.loc),
    vstatus: T(r.vstatus), seg: T(r.seg)
  })).filter(r => r.doc || r.item || r.agg);

  // ---- Target + IF (parsed; not displayed yet) ----
  const target = wsTarget ? readSheet(wsTarget, FIELDS.target).map(r => ({
    id: T(r.id), amt: N(r.amt), country: T(r.country), ctype: T(r.ctype), type: T(r.type)
  })).filter(r => r.id) : [];
  const ifs = wsIf ? readSheet(wsIf, FIELDS.ifs).map(r => ({
    iid: T(r.iid), sub: T(r.sub), d: toISO(r.date), type: T(r.type), doc: T(r.doc),
    name: T(r.name), status: T(r.status), amt: N(r.amt), cat: T(r.cat), rep: T(r.rep),
    loc: locNorm(r.loc), country: T(r.country), created: T(r.created)
  })) : [];

  // ---- Product Segment backfill (most-common segment per item / memo) ----
  const segByItem = new Map(), segByMemo = new Map();
  const bump = (map, key, seg) => {
    key = T(key); seg = T(seg); if (!key || !seg) return;
    if (!map.has(key)) map.set(key, new Map());
    const c = map.get(key); c.set(seg, (c.get(seg) || 0) + 1);
  };
  const best = (map, key) => {
    key = T(key); if (!key || !map.has(key)) return '';
    return [...map.get(key).entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])))[0][0];
  };
  sales.forEach(r => { if (r.seg) { bump(segByItem, r.item, r.seg); bump(segByMemo, r.memo, r.seg); } });
  sales.forEach(r => { if (!r.seg) r.seg = best(segByItem, r.item) || best(segByMemo, r.memo) || ''; });
  pff.forEach(r => { if (!r.seg) r.seg = best(segByItem, r.item) || best(segByMemo, r.desc) || ''; });

  return { sales, pff, target, ifs };
}

// grouped sum helper: rows -> [[label, value], ...] sorted desc, top n
function topBy(rows, keyFn, valFn, n) {
  const m = new Map();
  for (const r of rows) { const k = keyFn(r); if (k === '' || k == null) continue; m.set(k, (m.get(k) || 0) + valFn(r)); }
  const arr = [...m.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
  return n ? arr.slice(0, n) : arr;
}

function computeModel(data) {
  const { sales, pff, target, ifs } = data;

  // as-of date = latest sales date in the file; year = its year
  const asOf = sales.reduce((mx, r) => r.d > mx ? r.d : mx, '');
  const year = Number(asOf.slice(0, 4));
  const prev = year - 1;
  const yearStart = year + '-01-01';
  const ytdFrom = yearStart, ytdTo = asOf;
  const lyFrom = lyISO(ytdFrom), lyTo = lyISO(ytdTo);

  const inR = (d, a, b) => d && d >= a && d <= b;
  const selYTD = sales.filter(r => inR(r.d, ytdFrom, ytdTo));
  const selLY = sales.filter(r => inR(r.d, lyFrom, lyTo));

  const ytd = selYTD.reduce((a, r) => a + r.amt, 0);
  const ly = selLY.reduce((a, r) => a + r.amt, 0);
  const delta = ytd - ly, deltaPct = ly ? delta / ly : 0;
  const activeCustomers = new Set(selYTD.map(r => r.c).filter(Boolean)).size;
  const countries = new Set(selYTD.map(r => r.country).filter(Boolean)).size;

  // ---- monthly bars: complete months only (partial current month excluded) ----
  const asOfMonth = Number(asOf.slice(5, 7)), asOfDay = Number(asOf.slice(8, 10));
  const lastDayOfMonth = new Date(year, asOfMonth, 0).getDate();
  const monthComplete = asOfDay >= lastDayOfMonth;
  const lastFullMonth = monthComplete ? asOfMonth : asOfMonth - 1;
  const monthsIdx = [];
  for (let m = 1; m <= lastFullMonth; m++) monthsIdx.push(m);
  const monSum = (yr, mo) => sales.reduce((a, r) => a + (r.d.slice(0, 4) === String(yr) && Number(r.d.slice(5, 7)) === mo ? r.amt : 0), 0);
  const monthly = {
    idx: monthsIdx,
    cur: monthsIdx.map(m => monSum(year, m)),
    prev: monthsIdx.map(m => monSum(prev, m)),
    partialMonth: monthComplete ? null : asOfMonth  // 1-based index of excluded month, or null
  };

  // ---- markets (YTD) ----
  const topCountries = topBy(selYTD, r => r.country, r => r.amt, 7);
  const topSegments = topBy(selYTD, r => r.seg, r => r.amt, 6);
  const topCountriesShare = ytd ? topCountries.reduce((a, x) => a + x[1], 0) / ytd : 0;
  const topSegShare = ytd && topSegments.length ? topSegments[0][1] / ytd : 0;

  // ---- open order book (all pending-fulfillment rows) ----
  const openValue = pff.reduce((a, r) => a + r.agg, 0);
  const openOrders = new Set(pff.map(r => r.doc).filter(Boolean)).size;
  const bookCustomers = new Set(pff.map(r => r.c).filter(Boolean)).size;
  const readyValue = pff.reduce((a, r) => a + r.up * r.com, 0);
  const readyUnits = pff.reduce((a, r) => a + r.com, 0);
  const orderedUnits = pff.reduce((a, r) => a + r.ord, 0);
  const boRows = pff.filter(r => r.bo > 0);
  const boValue = boRows.reduce((a, r) => a + r.up * r.bo, 0);
  const boUnits = boRows.reduce((a, r) => a + r.bo, 0);
  const boSOs = new Set(boRows.map(r => r.doc).filter(Boolean)).size;

  // status: distinct SO (doc) count per validated status
  const statusOfDoc = new Map();
  pff.forEach(r => { if (r.doc && r.vstatus && !statusOfDoc.has(r.doc)) statusOfDoc.set(r.doc, r.vstatus); });
  const statusCount = new Map();
  for (const st of statusOfDoc.values()) statusCount.set(st, (statusCount.get(st) || 0) + 1);
  const status = [...statusCount.entries()].sort((a, b) => b[1] - a[1]);

  // ---- product / segment performance (YTD vs same-period-LY vs full prior year) ----
  const fyFrom = prev + '-01-01', fyTo = prev + '-12-31';
  const topLabel = (counts) => { let best = '', n = -1; for (const [k, c] of counts) if (c > n) { n = c; best = k; } return best; };

  // per-item aggregation across the three windows, in one pass over all sales
  const items = new Map();
  const ensureItem = it => { if (!items.has(it)) items.set(it, { item: it, ytd: 0, units: 0, ly: 0, fy: 0, bo: 0, boUnits: 0, _names: new Map(), _segs: new Map() }); return items.get(it); };
  sales.forEach(r => {
    if (!r.item) return;
    const o = ensureItem(r.item);
    if (r.memo) o._names.set(r.memo, (o._names.get(r.memo) || 0) + 1);
    if (r.seg) o._segs.set(r.seg, (o._segs.get(r.seg) || 0) + 1);
    if (inR(r.d, ytdFrom, ytdTo)) { o.ytd += r.amt; o.units += r.q; }
    if (inR(r.d, lyFrom, lyTo)) o.ly += r.amt;
    if (inR(r.d, fyFrom, fyTo)) o.fy += r.amt;
  });
  // fold in backorders by item (value = unit price × units on backorder)
  pff.filter(r => r.bo > 0 && r.item).forEach(r => {
    const o = ensureItem(r.item);
    o.bo += r.up * r.bo; o.boUnits += r.bo;
    if (r.desc) o._names.set(r.desc, (o._names.get(r.desc) || 0) + 1);
    if (r.seg) o._segs.set(r.seg, (o._segs.get(r.seg) || 0) + 1);
  });
  const prodPerf = [...items.values()].map(o => ({
    item: o.item, name: topLabel(o._names) || o.item, seg: topLabel(o._segs),
    ytd: o.ytd, units: o.units, ly: o.ly, fy: o.fy, bo: o.bo, boUnits: o.boUnits
  }));

  // per-segment aggregation (same windows + backorder)
  const segs = new Map();
  const ensureSeg = s => { if (!segs.has(s)) segs.set(s, { seg: s, ytd: 0, units: 0, ly: 0, fy: 0, bo: 0, boUnits: 0 }); return segs.get(s); };
  sales.forEach(r => {
    const s = r.seg || ''; const o = ensureSeg(s);
    if (inR(r.d, ytdFrom, ytdTo)) { o.ytd += r.amt; o.units += r.q; }
    if (inR(r.d, lyFrom, lyTo)) o.ly += r.amt;
    if (inR(r.d, fyFrom, fyTo)) o.fy += r.amt;
  });
  pff.filter(r => r.bo > 0).forEach(r => { const o = ensureSeg(r.seg || ''); o.bo += r.up * r.bo; o.boUnits += r.bo; });
  const segPerf = [...segs.values()];

  return {
    asOf, year, prev, yearStart,
    sales: { ytd, ly, delta, deltaPct, activeCustomers, countries, monthly },
    markets: { topCountries, topSegments, topCountriesShare, topSegShare },
    book: {
      openValue, openOrders, bookCustomers, readyValue, readyUnits, orderedUnits,
      boValue, boUnits, boSOs, status,
      boWh: topBy(boRows, r => r.loc, r => r.up * r.bo, 5),
      boCust: topBy(boRows, r => r.c, r => r.up * r.bo, 5),
      readyWh: topBy(pff, r => r.loc, r => r.up * r.com, 5),
      readyCust: topBy(pff, r => r.c, r => r.up * r.com, 5),
      readyShare: openValue ? readyValue / openValue : 0
    },
    perf: { prodPerf, segPerf },
    counts: { sales: sales.length, pff: pff.length, target: target.length, ifs: ifs.length }
  };
}

/* ------------------------------------------------------------------ *
 * 4. Charts (SVG; read CSS vars so they follow the active theme)
 * ------------------------------------------------------------------ */
const css = k => getComputedStyle(document.documentElement).getPropertyValue(k).trim();
const ttEl = () => $('tt');
function wireTips(sel) {
  document.querySelectorAll(sel).forEach(el => {
    el.addEventListener('mousemove', e => { const t = ttEl(); t.innerHTML = el.getAttribute('data-tip'); t.style.opacity = 1; t.style.left = (e.clientX + 13) + 'px'; t.style.top = (e.clientY - 12) + 'px'; });
    el.addEventListener('mouseleave', () => { ttEl().style.opacity = 0; });
  });
}
const monLabels = () => I18N.MONTHS[LANG];

// toggle the left/right edge-fade cues on each scrollable perf table
function updateScrollShadows() {
  document.querySelectorAll('.ptable-scroll').forEach(sc => {
    const w = sc.querySelector('.ptable-wrap'); if (!w) return;
    const max = w.scrollWidth - w.clientWidth;
    sc.classList.toggle('more-right', w.scrollLeft < max - 1);
    sc.classList.toggle('more-left', w.scrollLeft > 1);
  });
}

// ---- FLIP re-rank animation for the performance tables ----
const reduceMotion = () => window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
// snapshot each keyed row's viewport position before a re-render
function captureRowRects() {
  const map = new Map();
  document.querySelectorAll('.ptable tbody tr[data-key]').forEach(tr => map.set(tr.dataset.key, tr.getBoundingClientRect().top));
  return map;
}
// after the re-render, tween each row from its old position to its new one.
// Uses the Web Animations API (fill: none) so no inline styles linger — a row
// can never get stuck offset or invisible, even under rapid toggling.
function flipRerank(prev) {
  if (!prev || reduceMotion() || !document.body.animate) return;
  const rows = [...document.querySelectorAll('.ptable tbody tr[data-key]')];
  const ease = 'cubic-bezier(.22,.61,.36,1)';
  rows.forEach((tr, i) => {
    const key = tr.dataset.key, delay = Math.min(i * 10, 180);
    if (prev.has(key)) {
      const dy = prev.get(key) - tr.getBoundingClientRect().top;   // First - Last
      if (Math.abs(dy) > 0.5) tr.animate(
        [{ transform: `translateY(${dy}px)` }, { transform: 'translateY(0)' }],
        { duration: 550, delay, easing: ease });
    } else {                                                        // row entering the ranking
      tr.animate(
        [{ opacity: 0, transform: 'translateY(6px)' }, { opacity: 1, transform: 'translateY(0)' }],
        { duration: 450, delay, easing: 'ease' });
    }
  });
}

function monthlyBarsSVG(mon) {
  const labels = mon.idx.map(i => monLabels()[i - 1]);
  const cur = mon.cur, prv = mon.prev, n = labels.length;
  const W = 1080, H = 300, padL = 58, padR = 14, padT = 14, padB = 30, iw = W - padL - padR, ih = H - padT - padB;
  const max = Math.max(1, ...cur, ...prv), ymax = Math.ceil(max / 2e5) * 2e5 || 2e5;
  const y = v => padT + ih - (v / ymax) * ih, gw = iw / (n || 1), bw = Math.min(38, gw * 0.32), gap = 7;
  let s = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img">`;
  for (let v = 0; v <= ymax; v += 2e5) { const yy = y(v); s += `<line x1="${padL}" y1="${yy}" x2="${W - padR}" y2="${yy}" stroke="${css('--grid')}"/>`; let lab = (v / 1e6).toFixed(1); if (LANG === 'es') lab = lab.replace('.', ','); s += `<text x="${padL - 9}" y="${yy + 4}" text-anchor="end" font-size="11.5" fill="${css('--muted')}">$${lab}M</text>`; }
  labels.forEach((m, i) => {
    const gx = padL + gw * i + gw / 2;
    [[gx - bw - gap / 2, cur[i], '--cur', mon.yearCur], [gx + gap / 2, prv[i], '--prev', mon.yearPrev]].forEach(([bx, val, c, yr]) => {
      const by = y(val), bh = padT + ih - by;
      s += `<rect x="${bx}" y="${by}" width="${bw}" height="${Math.max(0, bh)}" rx="4" fill="${css(c)}" data-tip="${yr} ${m} · ${full(val)}"/>`;
    });
    s += `<text x="${gx}" y="${H - 9}" text-anchor="middle" font-size="12" fill="${css('--ink2')}">${m}</text>`;
  });
  s += `<line x1="${padL}" y1="${padT + ih}" x2="${W - padR}" y2="${padT + ih}" stroke="${css('--axis')}" stroke-width="1.5"/></svg>`;
  return s;
}

function cumulativeSVG(mon) {
  const labels = mon.idx.map(i => monLabels()[i - 1]);
  const cum = a => { let t = 0; return a.map(v => t += v); };
  const c26 = cum(mon.cur), c25 = cum(mon.prev), n = labels.length;
  const W = 1080, H = 300, padL = 58, padR = 18, padT = 16, padB = 30, iw = W - padL - padR, ih = H - padT - padB;
  const max = Math.max(1, ...c26, ...c25), ymax = Math.ceil(max / 1e6) * 1e6 || 1e6;
  const x = i => padL + iw * (n > 1 ? i / (n - 1) : 0.5), y = v => padT + ih - (v / ymax) * ih;
  let s = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img">`;
  for (let v = 0; v <= ymax; v += 1e6) { const yy = y(v); s += `<line x1="${padL}" y1="${yy}" x2="${W - padR}" y2="${yy}" stroke="${css('--grid')}"/>`; s += `<text x="${padL - 9}" y="${yy + 4}" text-anchor="end" font-size="11.5" fill="${css('--muted')}">$${v / 1e6}M</text>`; }
  const pts26 = c26.map((v, i) => x(i) + ',' + y(v)).join(' '), pts25 = c25.map((v, i) => x(i) + ',' + y(v)).join(' ');
  if (n > 0) s += `<polygon points="${padL},${padT + ih} ${pts26} ${W - padR},${padT + ih}" fill="${css('--cur')}" opacity="0.07"/>`;
  s += `<polyline points="${pts25}" fill="none" stroke="${css('--prev')}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>`;
  s += `<polyline points="${pts26}" fill="none" stroke="${css('--cur')}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>`;
  [[c25, '--prev', mon.yearPrev], [c26, '--cur', mon.yearCur]].forEach(([arr, cv, yr]) => arr.forEach((v, i) => {
    s += `<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="${css(cv)}" stroke="${css('--surface')}" stroke-width="1.5" data-tip="${yr} ${labels[i]} · ${full(v)}"/>`;
  }));
  labels.forEach((m, i) => { s += `<text x="${x(i)}" y="${H - 9}" text-anchor="middle" font-size="12" fill="${css('--ink2')}">${m}</text>`; });
  s += `<line x1="${padL}" y1="${padT + ih}" x2="${W - padR}" y2="${padT + ih}" stroke="${css('--axis')}" stroke-width="1.5"/></svg>`;
  return s;
}

function hbarsHTML(rows, colorVar) {
  if (!rows.length) return '<div class="note">—</div>';
  const max = Math.max(...rows.map(r => r[1])) || 1, c = css(colorVar);
  return rows.map(([name, val]) =>
    `<div class="hbar"><span class="name" title="${esc(name)}">${esc(name)}</span>` +
    `<span class="track"><span class="fill" style="width:${(val / max * 100).toFixed(1)}%;background:${c}"></span></span>` +
    `<span class="amt">${short(val)}</span></div>`).join('');
}
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m])); }

// delta cell: signed money, colored by direction
const dCls = n => 'num ' + (n > 0 ? 'good' : n < 0 ? 'bad' : '');
const deltaCell = n => ({ v: signed(n), cls: dCls(n) });
const pctCell = (x, ref) => ({ v: (x > 0 ? '+' : '') + pct1(x), cls: dCls(ref) });

// generic performance table (actual or backorder-adjusted)
function perfTable(isProduct, mode, list, limit) {
  const str = S(), bo = mode === 'bo', prev = MODEL.prev;
  const val = r => bo ? r.ytd + r.bo : r.ytd;              // sort / comparison base
  const rows = list.slice().sort((a, b) => val(b) - val(a)).slice(0, limit || list.length);
  if (!rows.length) return `<div class="note">${str.noItems}</div>`;

  const H = [{ label: '#', num: true, tip: str.tipRank },
             { label: isProduct ? str.colProduct : str.colSegment, tip: isProduct ? str.tipProduct : str.tipSegment }];
  if (isProduct) H.push({ label: str.colSegment, tip: str.tipSegment });
  H.push({ label: str.colYTD, num: true, tip: str.tipYTD });
  if (bo) { H.push({ label: str.colBO, num: true, tip: str.tipBO }); H.push({ label: str.colAdj, num: true, tip: str.tipAdj }); }
  else H.push({ label: str.colUnits, num: true, tip: str.tipUnits });
  H.push({ label: fmt(str.colSP, { prev }), num: true, tip: fmt(str.tipSP, { prev }) }, { label: str.colDSP, num: true, tip: fmt(str.tipDSP, { prev }) }, { label: str.colPSP, num: true, tip: fmt(str.tipPSP, { prev }) });
  H.push({ label: fmt(str.colFY, { prev }), num: true, tip: fmt(str.tipFY, { prev }) }, { label: str.colDFY, num: true, tip: fmt(str.tipDFY, { prev }) }, { label: str.colPFY, num: true, tip: fmt(str.tipPFY, { prev }) });

  const body = rows.map((r, i) => {
    const base = val(r), dsp = base - r.ly, dfy = base - r.fy;
    const key = (isProduct ? 'p:' + r.item : 's:' + r.seg);
    const c = [{ v: i + 1, cls: 'rank' }];
    if (isProduct) c.push({ html: `<b>${esc(r.item)}</b><span>${esc(r.name)}</span>`, cls: 'name' });
    else c.push({ v: r.seg ? I18N.seg(r.seg, LANG) : 'N/A', cls: 'strong' });
    if (isProduct) c.push({ v: r.seg ? I18N.seg(r.seg, LANG) : 'N/A' });
    c.push({ v: full(r.ytd), cls: 'num' });
    if (bo) { c.push({ v: r.bo ? full(r.bo) : '—', cls: 'num' + (r.bo ? ' bo-add' : '') }); c.push({ v: full(base), cls: 'num strong' }); }
    else c.push({ v: num(r.units), cls: 'num' });
    c.push({ v: full(r.ly), cls: 'num' }, deltaCell(dsp), pctCell(r.ly ? dsp / r.ly : 0, dsp));
    c.push({ v: full(r.fy), cls: 'num' }, deltaCell(dfy), pctCell(r.fy ? dfy / r.fy : 0, dfy));
    return { key, cells: c };
  });

  return `<div class="ptable-scroll"><div class="ptable-wrap"><table class="ptable"><thead><tr>${H.map(h => `<th class="${h.num ? 'num' : ''}"${h.tip ? ` data-tip="${esc(h.tip)}"` : ''}>${esc(h.label)}</th>`).join('')}</tr></thead><tbody>${body.map(row => `<tr data-key="${esc(row.key)}">${row.cells.map(x => `<td class="${x.cls || ''}">${x.html || esc(x.v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
}

// top-N segmented control for the product table
function limitTools(cur) {
  const str = S();
  return `<div class="table-tools"><span class="lbl">${str.showN}</span><div class="seg">` +
    [10, 25, 50, 'all'].map(v => {
      const target = v === 'all' ? Infinity : v, label = v === 'all' ? str.showAll : v;
      return `<button class="seg-btn ${cur === target ? 'active' : ''}" data-plimit="${v}">${label}</button>`;
    }).join('') + `</div></div>`;
}

// per-table view toggle: actual sales vs. backorder-adjusted (what-if-fulfilled)
function viewTools(kind) {
  const str = S(), on = kind === 'seg' ? BO_SEG : BO_PROD, attr = 'data-bo' + kind;
  return `<div class="table-tools perf-view"><span class="lbl">${str.perfView}</span><div class="seg">` +
    `<button class="seg-btn ${!on ? 'active' : ''}" ${attr}="0">${str.perfActual}</button>` +
    `<button class="seg-btn ${on ? 'active' : ''}" ${attr}="1">${str.perfBO}</button>` +
    `</div></div>`;
}

/* ------------------------------------------------------------------ *
 * 5. Render
 * ------------------------------------------------------------------ */
function render() {
  document.documentElement.setAttribute('lang', LANG);
  document.documentElement.setAttribute('data-theme', THEME);
  const str = S();

  // header
  $('t-title').textContent = str.title;
  $('t-subtitle').textContent = str.subtitle;
  $('btn-reupload').textContent = str.reupload;
  $('btn-share').textContent = str.share;
  $('loading').hidden = true;

  // language / theme button states
  document.querySelectorAll('#lang-seg .seg-btn').forEach(b => b.classList.toggle('active', b.dataset.lang === LANG));
  document.querySelectorAll('#theme-seg .seg-btn').forEach(b => b.classList.toggle('active', b.dataset.themeBtn === THEME));

  if (!MODEL) { renderEmpty(); return; }

  const m = MODEL, sd = m.sales, bk = m.book, mk = m.markets, pf = m.perf;

  // as-of box
  $('asof-box').hidden = false;
  $('asof-big').textContent = fmt(str.asOfBig, { date: longDate(m.asOf) });
  $('asof-mut').textContent = fmt(str.asOfMut, { start: shortDM(m.yearStart), date: shortDM(m.asOf) });
  $('btn-reupload').hidden = false;
  $('btn-share').hidden = false;

  // month names for charts
  sd.monthly.yearCur = m.year; sd.monthly.yearPrev = m.prev;
  const partialName = sd.monthly.partialMonth ? I18N.MONTHS_LONG[LANG][sd.monthly.partialMonth - 1] : '';

  const arrow = sd.delta >= 0 ? '▲' : '▼';
  const pillCls = sd.delta >= 0 ? 'up' : 'down';
  const pillTxt = fmt(str.pillVs, {
    arrow, sign: sd.delta >= 0 ? '+' : '', pct: pct1(Math.abs(sd.deltaPct)),
    dsign: sd.delta >= 0 ? '+' : '', delta: short(sd.delta), prev: m.prev
  });

  const legMonthly = sd.monthly.partialMonth
    ? fmt(str.legMonthly, { month: partialName })
    : (LANG === 'es' ? 'Ventas netas mensuales' : 'Monthly net sales');

  // status chips + units chip
  const chips = bk.status.map(([k, v]) => `<span class="chip">${esc(I18N.status(k, LANG))} <b>${num(v)}</b></span>`).join('') +
    `<span class="chip">${str.s3units} <b>${num(bk.orderedUnits)}</b></span>`;

  // translate market rows
  const ctryRows = mk.topCountries.map(([k, v]) => [I18N.country(k, LANG), v]);
  const segRows = mk.topSegments.map(([k, v]) => [I18N.seg(k, LANG), v]);
  const whTr = rows => rows.map(([k, v]) => [I18N.wh(k, LANG), v]);
  const topSegName = mk.topSegments.length ? I18N.seg(mk.topSegments[0][0], LANG) : '—';

  // per-table view mode + abbreviation line (actual vs. backorder-adjusted)
  const segMode = BO_SEG ? 'bo' : 'actual', prodMode = BO_PROD ? 'bo' : 'actual';
  const abbrActual = fmt(str.abbrPerf, { prev: m.prev, year: m.year });
  const segAbbr = BO_SEG ? str.abbrPerfBO : abbrActual;
  const prodAbbr = BO_PROD ? str.abbrPerfBO : abbrActual;

  $('dashboard').hidden = false;
  $('dashboard').innerHTML = `
  <!-- Section 1 -->
  <div class="section">
    <div class="shead"><div class="snum">1</div><div><h2>${str.s1h}</h2><div class="st">${fmt(str.s1st, { prev: m.prev })}</div></div></div>
    <div class="row r-311">
      <div class="card">
        <div class="label">${fmt(str.s1curLabel, { year: m.year })}</div>
        <div class="hero">${short(sd.ytd)}</div>
        <div class="pill ${pillCls}">${pillTxt}</div>
        <div class="note">${fmt(str.s1curNote, { start: shortDM(m.yearStart), date: shortDM(m.asOf), year: m.year, full: full(sd.ytd) })}</div>
      </div>
      <div class="card">
        <div class="label">${fmt(str.s1prevLabel, { prev: m.prev })}</div>
        <div class="value">${short(sd.ly)}</div>
        <div class="note">${fmt(str.s1prevNote, { start: shortDM(lyISO(m.yearStart)), date: shortDM(lyISO(m.asOf)), prev: m.prev })}</div>
      </div>
      <div class="card">
        <div class="label">${str.s1custLabel}</div>
        <div class="value">${num(sd.activeCustomers)}</div>
        <div class="note">${fmt(str.s1custNote, { n: sd.countries })}</div>
      </div>
    </div>
    <div class="card mt16">
      <div class="legend">
        <span><span class="sw" style="background:var(--cur)"></span>${m.year}</span>
        <span><span class="sw" style="background:var(--prev)"></span>${m.prev}</span>
        <span style="color:var(--muted)">${legMonthly}</span>
      </div>
      <div id="chart-trend"></div>
    </div>
    <div class="card mt16">
      <div class="legend">
        <span><span class="sw" style="background:var(--cur)"></span>${fmt(str.legCumCur, { year: m.year })}</span>
        <span><span class="sw" style="background:var(--prev)"></span>${fmt(str.legCumPrev, { prev: m.prev })}</span>
        <span style="color:var(--muted)">${str.legCumNote}</span>
      </div>
      <div id="chart-cum"></div>
    </div>
  </div>

  <!-- Section 2 -->
  <div class="section">
    <div class="shead"><div class="snum">2</div><div><h2>${str.s2h}</h2><div class="st">${fmt(str.s2st, { year: m.year })}</div></div></div>
    <div class="row r-2">
      <div class="card">
        <div class="label">${str.s2ctryLabel}</div>
        <div class="note" style="margin:2px 0 14px">${fmt(str.s2ctryNote, { pct: pct0(mk.topCountriesShare) })}</div>
        ${hbarsHTML(ctryRows, '--cur')}
      </div>
      <div class="card">
        <div class="label">${str.s2segLabel}</div>
        <div class="note" style="margin:2px 0 14px">${fmt(str.s2segNote, { top: esc(topSegName), pct: pct0(mk.topSegShare) })}</div>
        ${hbarsHTML(segRows, '--ready')}
      </div>
    </div>
  </div>

  <!-- Section 3 -->
  <div class="section">
    <div class="shead"><div class="snum">3</div><div><h2>${str.s3h}</h2><div class="st">${str.s3st}</div></div></div>
    <div class="card accent a-open">
      <div class="row r-311" style="align-items:center">
        <div>
          <div class="label">${str.s3valLabel}</div>
          <div class="hero">${short(bk.openValue)}</div>
          <div class="note">${fmt(str.s3valNote, { n: num(bk.openOrders), m: num(bk.bookCustomers) })}</div>
        </div>
        <div style="grid-column:span 2">
          <div class="label" style="margin-bottom:9px">${str.s3compLabel}</div>
          <div class="comp" id="comp"></div>
          <div class="comp-key">
            <span class="k"><span class="sw" style="display:inline-block;width:11px;height:11px;border-radius:3px;background:var(--ready);margin-right:7px;vertical-align:-1px"></span>${str.s3ready} <b>${short(bk.readyValue)}</b></span>
            <span class="k"><span class="sw" style="display:inline-block;width:11px;height:11px;border-radius:3px;background:var(--bo);margin-right:7px;vertical-align:-1px"></span>${str.s3bo} <b>${short(bk.boValue)}</b></span>
          </div>
          <div class="status-pills">${chips}</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Sections 4 & 5 -->
  <div class="row r-2 section">
    <div>
      <div class="shead"><div class="snum">4</div><div><h2>${str.s4h}</h2><div class="st">${str.s4st}</div></div></div>
      <div class="card accent a-bo" style="height:calc(100% - 54px)">
        <div class="label">${str.s4valLabel}</div>
        <div class="hero" style="color:var(--bo)">${short(bk.boValue)}</div>
        <div class="subgrid" style="margin-top:14px">
          <div><div class="label">${str.s4units}</div><div class="value" style="font-size:22px">${num(bk.boUnits)}</div></div>
          <div><div class="label">${str.s4sos}</div><div class="value" style="font-size:22px">${num(bk.boSOs)}</div></div>
        </div>
        <div class="mini-t">${str.byWh}</div>${hbarsHTML(whTr(bk.boWh), '--bo')}
        <div class="mini-t">${str.byCust}</div>${hbarsHTML(bk.boCust, '--bo')}
      </div>
    </div>
    <div>
      <div class="shead"><div class="snum">5</div><div><h2>${str.s5h}</h2><div class="st">${str.s5st}</div></div></div>
      <div class="card accent a-ready" style="height:calc(100% - 54px)">
        <div class="label">${str.s5valLabel}</div>
        <div class="hero" style="color:var(--ready)">${short(bk.readyValue)}</div>
        <div class="subgrid" style="margin-top:14px">
          <div><div class="label">${str.s5unitsLabel}</div><div class="value" style="font-size:22px">${num(bk.readyUnits)}</div></div>
          <div><div class="label">${str.s5shareLabel}</div><div class="value" style="font-size:22px">${pct0(bk.readyShare)}</div></div>
        </div>
        <div class="mini-t">${str.byWh}</div>${hbarsHTML(whTr(bk.readyWh), '--ready')}
        <div class="mini-t">${str.byCust}</div>${hbarsHTML(bk.readyCust, '--ready')}
      </div>
    </div>
  </div>

  <!-- Section 6 — product performance (each table toggles actual ⇄ backorder-adjusted) -->
  <div class="section">
    <div class="shead"><div class="snum">6</div><div><h2>${str.s6h}</h2><div class="st">${str.s6st}</div></div></div>
    <div class="card">
      <div class="card-title">${str.tblSeg}</div>
      <div class="perf-ctrls">${viewTools('seg')}</div>
      <div class="abbr">${segAbbr}</div>
      ${perfTable(false, segMode, pf.segPerf)}
    </div>
    <div class="card mt16">
      <div class="card-title">${str.tblProd}</div>
      <div class="perf-ctrls">${viewTools('prod')}${limitTools(ITEM_LIMIT)}</div>
      <div class="abbr">${prodAbbr}</div>
      ${perfTable(true, prodMode, pf.prodPerf, ITEM_LIMIT)}
    </div>
  </div>`;

  // charts + composition bar (after DOM insert; they read theme CSS vars)
  $('chart-trend').innerHTML = monthlyBarsSVG(sd.monthly);
  $('chart-cum').innerHTML = cumulativeSVG(sd.monthly);
  const tot = bk.readyValue + bk.boValue || 1, rp = bk.readyValue / tot * 100, bp = bk.boValue / tot * 100;
  $('comp').innerHTML =
    `<span style="flex:${rp};background:var(--ready);color:var(--on-neutral)">${Math.round(rp)}%</span>` +
    `<span style="flex:${bp};background:var(--bo);color:#fff">${Math.round(bp)}%</span>`;
  wireTips('#chart-trend rect'); wireTips('#chart-cum circle'); wireTips('.ptable th[data-tip]');

  // per-table view toggles (animated re-rank) + product top-N
  document.querySelectorAll('[data-boseg]').forEach(b => b.onclick = () => {
    const next = b.dataset.boseg === '1'; if (next === BO_SEG) return;
    const prev = captureRowRects(); BO_SEG = next; render(); flipRerank(prev);
  });
  document.querySelectorAll('[data-boprod]').forEach(b => b.onclick = () => {
    const next = b.dataset.boprod === '1'; if (next === BO_PROD) return;
    const prev = captureRowRects(); BO_PROD = next; render(); flipRerank(prev);
  });
  document.querySelectorAll('[data-plimit]').forEach(b => b.onclick = () => { ITEM_LIMIT = b.dataset.plimit === 'all' ? Infinity : Number(b.dataset.plimit); render(); });

  // horizontal-scroll edge cues for the (wide) performance tables
  document.querySelectorAll('.ptable-wrap').forEach(w => w.addEventListener('scroll', updateScrollShadows, { passive: true }));
  updateScrollShadows();

  // footer
  $('foot').hidden = false;
  $('foot').innerHTML = fmt(str.foot, { source: esc(SOURCE || '—'), date: longDate(m.asOf) });
}

function renderEmpty() {
  $('asof-box').hidden = true;
  $('btn-reupload').hidden = true;
  $('btn-share').hidden = true;
  $('loading').hidden = true;
  $('dashboard').hidden = true;
  $('foot').hidden = true;
  $('uploader').hidden = false;
  const str = S();
  $('up-title').textContent = str.upTitle;
  $('up-sub').innerHTML = str.upSub;
  $('btn-browse').textContent = str.upBrowse;
  $('up-fmt').innerHTML = str.upFmt;
  $('up-sample-text').textContent = str.sampleText;
  $('btn-sample').textContent = str.sampleLink;
}

/* ------------------------------------------------------------------ *
 * 6. Shareable link — embed the computed model in the URL hash.
 *    No backend: the model (not the raw rows) is compacted, gzipped
 *    (CompressionStream) and base64url-encoded into '#d='. The hash is
 *    never sent to the server, so the data stays client-side.
 * ------------------------------------------------------------------ */
function bytesToB64url(bytes) {
  let bin = ''; const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function b64urlToBytes(s) {
  s = s.replace(/-/g, '+').replace(/_/g, '/'); while (s.length % 4) s += '=';
  const bin = atob(s), out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
async function gzipBytes(str) { const s = new Blob([str]).stream().pipeThrough(new CompressionStream('gzip')); return new Uint8Array(await new Response(s).arrayBuffer()); }
async function gunzipStr(bytes) { const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip')); return new TextDecoder().decode(await new Response(s).arrayBuffer()); }
async function encodeShare(str) {
  if (typeof CompressionStream !== 'undefined') return 'g' + bytesToB64url(await gzipBytes(str));
  return 'r' + bytesToB64url(new TextEncoder().encode(str));   // fallback: uncompressed
}
async function decodeShare(enc) {
  const fmt = enc[0], bytes = b64urlToBytes(enc.slice(1));
  if (fmt === 'g') return await gunzipStr(bytes);
  if (fmt === 'r') return new TextDecoder().decode(bytes);
  return await gunzipStr(b64urlToBytes(enc));                  // legacy: no prefix
}

// compact the two large arrays (objects -> positional arrays, rounded) and back
const _R = x => Math.round(Number(x) || 0);
function compactModel(m) {
  const c = JSON.parse(JSON.stringify(m));
  c.perf.prodPerf = m.perf.prodPerf.map(r => [r.item, r.name, r.seg, _R(r.ytd), _R(r.units), _R(r.ly), _R(r.fy), _R(r.bo), _R(r.boUnits)]);
  c.perf.segPerf = m.perf.segPerf.map(r => [r.seg, _R(r.ytd), _R(r.units), _R(r.ly), _R(r.fy), _R(r.bo), _R(r.boUnits)]);
  return c;
}
function hydrateModel(c) {
  c.perf.prodPerf = c.perf.prodPerf.map(a => ({ item: a[0], name: a[1], seg: a[2], ytd: a[3], units: a[4], ly: a[5], fy: a[6], bo: a[7], boUnits: a[8] }));
  c.perf.segPerf = c.perf.segPerf.map(a => ({ seg: a[0], ytd: a[1], units: a[2], ly: a[3], fy: a[4], bo: a[5], boUnits: a[6] }));
  return c;
}

// copy to clipboard, but never hang: if the async Clipboard API doesn't
// settle quickly (denied permission / unfocused page), fall back to execCommand.
function copyText(text) {
  return new Promise(resolve => {
    let done = false;
    const fallback = () => {
      if (done) return; done = true;
      let ok = false;
      try {
        const ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        ok = document.execCommand('copy'); document.body.removeChild(ta);
      } catch (e) { ok = false; }
      resolve(ok);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      const t = setTimeout(fallback, 700);
      navigator.clipboard.writeText(text).then(
        () => { if (done) return; done = true; clearTimeout(t); resolve(true); },
        () => { clearTimeout(t); fallback(); }
      );
    } else fallback();
  });
}

async function shareCurrent() {
  if (!MODEL) return;
  const btn = $('btn-share');
  btn.disabled = true; btn.textContent = S().shareBuilding;
  try {
    const payload = { v: 2, source: SOURCE, model: compactModel(MODEL) };
    const enc = await encodeShare(JSON.stringify(payload));
    const url = location.origin + location.pathname + '#d=' + enc;
    await copyText(url);
    btn.textContent = S().shareCopied;
  } catch (e) { console.error('share failed', e); btn.textContent = S().shareErr; }
  setTimeout(() => { btn.disabled = false; btn.textContent = S().share; }, 1900);
}

// ---- loading indicator (progress bar) ----
// Wait for the browser to paint before running blocking work. rAF gives us the
// paint, but it's PAUSED in background/unfocused tabs — so race it against a
// timeout fallback, otherwise the parse would hang if the tab isn't visible.
const yieldPaint = () => new Promise(r => {
  let done = false; const fin = () => { if (!done) { done = true; r(); } };
  requestAnimationFrame(() => requestAnimationFrame(fin));
  setTimeout(fin, 60);
});
function showLoading(label) {
  $('uploader').hidden = true; $('dashboard').hidden = true; $('foot').hidden = true;
  $('loading').hidden = false; setProgress(0, label);
}
function setProgress(pct, label) {
  const f = $('progress-fill'); if (f) f.style.width = Math.max(0, Math.min(100, pct)) + '%';
  if (label != null) { const l = $('loading-label'); if (l) l.textContent = label; }
}

async function loadShared(enc) {
  showLoading(S().loadingShared); setProgress(30);
  try {
    const payload = JSON.parse(await decodeShare(enc));
    setProgress(70);
    MODEL = hydrateModel(payload.model);
    SOURCE = payload.source || '';
    setProgress(100); await yieldPaint();
    render();
    window.scrollTo(0, 0);
  } catch (e) {
    console.error('shared link decode failed', e);
    MODEL = null; render(); showError(S().errShared);
  }
}

/* ------------------------------------------------------------------ *
 * 7. Sample workbook — dummy (obfuscated) data in the exact format,
 *    generated in the browser with SheetJS so users see what to upload.
 * ------------------------------------------------------------------ */
function buildSampleWorkbook() {
  const REPS = ['Alex Rivera', 'Jordan Kim', 'Sam Okafor', 'Chris Dubois'];
  const WH = ['MXWHS', 'EU Warehouse', 'ProWHS', 'VTNWH-MX'];
  const CUSTS = [
    { id: 'I/9001', name: 'I/9001 ACME MOTO DISTRIBUTION', country: 'Australia', ctype: 'EX-ASIA', cat: 'International Distributor' },
    { id: 'I/9002', name: 'I/9002 GLOBEX POWERSPORTS', country: 'Chile', ctype: 'EX-LATAM', cat: 'International Distributor' },
    { id: 'I/9003', name: 'I/9003 INITECH LUBRICANTS', country: 'Mexico', ctype: 'EX-LATAM', cat: 'International Distributor' },
    { id: 'I/9004', name: 'I/9004 UMBRELLA MOTORSPORT', country: 'Germany', ctype: 'EX-EU', cat: 'International Distributor' },
    { id: 'I/9005', name: 'I/9005 HOOLI RACING', country: 'Japan', ctype: 'EX-ASIA', cat: 'International Distributor' },
    { id: 'I/9006', name: 'I/9006 VEHEMENT DISTRIBUTION', country: 'Brazil', ctype: 'EX-LATAM', cat: 'International Distributor' }
  ];
  const ITEMS = [
    { sku: 'AA-1001', memo: 'SYNTHETIC 4T 10W40 / 1 LTR', seg: '4T Engine Oils', up: 38 },
    { sku: 'AA-2002', memo: '2T RACING PREMIX / 1 LTR', seg: '2T Engine Oils', up: 30 },
    { sku: 'AE-3003', memo: 'CHAIN LUBE AEROSOL / 400 ML', seg: 'Aerosols', up: 22 },
    { sku: 'SU-4004', memo: 'FORK OIL 5W / 1 LTR', seg: 'Suspension Oils', up: 41 },
    { sku: 'MN-5005', memo: 'CONTACT CLEANER / 500 ML', seg: 'Maintenance', up: 18 },
    { sku: 'TG-6006', memo: 'GEAR OIL 80W90 / 1 LTR', seg: 'Trans & Gear Oils', up: 34 }
  ];
  const pad = n => String(n).padStart(2, '0');

  const salesHead = ['ID', 'Customer', 'Sales Rep', 'Transaction Type', 'Date', 'Document Number', 'Item', 'Memo', 'Subsidiary: Name', 'Customer Category: Name', 'Quantity', 'Unit Price', 'Amount (Foreign Currency)', 'Amount', 'Customer/Project: Customer Type', 'Inventory Location: Name', 'Address: Billing Address Country Name', 'Product Segment', 'SO Date'];
  const sales = [salesHead];
  let doc = 50000;
  const addSale = (y, m, d, ci, ii, qty) => {
    const c = CUSTS[ci % CUSTS.length], it = ITEMS[ii % ITEMS.length], rep = REPS[(ci + ii) % REPS.length];
    const amt = Math.round(qty * it.up * 100) / 100;
    sales.push([c.id, c.name, rep, 'Invoice', `${y}-${pad(m)}-${pad(d)}`, 'INV' + (doc++), it.sku, it.memo,
      'Maxima Racing Oils', c.cat, qty, it.up, amt, amt, c.ctype, WH[ci % WH.length], c.country, it.seg, '']);
  };
  for (let m = 1; m <= 12; m++) for (let ci = 0; ci < CUSTS.length; ci++) for (let ii = 0; ii < ITEMS.length; ii++)
    if ((m + ci + ii) % 2 === 0) addSale(2025, m, 15, ci, ii, 300 + ((m * 37 + ci * 53 + ii * 71) % 900));
  for (let m = 1; m <= 6; m++) for (let ci = 0; ci < CUSTS.length; ci++) for (let ii = 0; ii < ITEMS.length; ii++)
    if ((m + ci + ii) % 2 === 1) addSale(2026, m, 10, ci, ii, 350 + ((m * 41 + ci * 47 + ii * 59) % 1000));
  addSale(2026, 7, 1, 0, 0, 800); addSale(2026, 7, 2, 2, 3, 600);   // partial current month

  const target = [['ID', 'Amount', 'Country', 'Customer Type', 'Type']];
  CUSTS.forEach((c, i) => target.push([c.id, 100000 + i * 25000, c.country, c.ctype, 'Target']));

  const pffHead = ['Item Type', 'Item', 'Description (Sales)', 'Date', 'Transaction Type', 'Primary Sales Rep', 'Document Number', 'Customer', 'Ordered', 'Fulfilled', 'Committed', 'Back Ordered', 'Subsidiary: Name', 'Aggregate Amount', 'Unit Price', 'Location', 'Validated Status'];
  const pff = [pffHead];
  let so = 70000;
  CUSTS.forEach((c, ci) => ITEMS.forEach((it, ii) => {
    if ((ci + ii) % 3 !== 0) return;
    const ord = 500 + ((ci * 137 + ii * 211) % 2500), bo = ii % 2 === 0 ? Math.round(ord * 0.3) : 0, com = ord - bo;
    pff.push(['Assembly', it.sku, it.memo, `2026-${pad((ii % 6) + 1)}-20`, 'Sales Order', REPS[(ci + ii) % REPS.length], 'SO' + (so++),
      c.name, ord, 0, com, bo, 'Maxima Racing Oils', Math.round(ord * it.up * 100) / 100, it.up, WH[ci % WH.length],
      bo > 0 ? 'Partially Fulfilled' : 'Pending Fulfillment', it.seg]);
  }));

  const ifHead = ['Internal ID', 'Subsidiary', 'Date', 'Type', 'Document Number', 'Transaction Number', 'Name', 'PO/Check Number', 'Status', 'Memo', 'Amount (Foreign Currency)', 'Amount', 'Category', 'Sales Rep', 'Location', 'Shipping Country', 'Terms', 'Created'];
  const ifs = [ifHead];
  for (let i = 0; i < 6; i++) ifs.push([2700000 + i, 'MRO Holding : Maxima Racing Oils', `2026-06-${pad(10 + i)}`, 'Item Fulfillment', 'IF' + (46000 + i),
    'ITEMSHIP' + (46000 + i), CUSTS[i].name, '', 'Released', 'Sample fulfillment', 0, 0, 'Retail/Web', REPS[i % REPS.length], 'MXWHS', 'United States', '', 'SO' + (70000 + i)]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sales), 'Sales');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(target), 'Target');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pff), 'PendingFullFill');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ifs), 'IF');
  return wb;
}
function downloadSample() {
  try { XLSX.writeFile(buildSampleWorkbook(), 'MRO_sample_data.xlsx'); }
  catch (e) { console.error('sample build failed', e); }
}

/* ------------------------------------------------------------------ *
 * 8. File handling
 * ------------------------------------------------------------------ */
function showError(msg) { const e = $('up-err'); e.hidden = false; e.textContent = msg; }
function clearError() { $('up-err').hidden = true; }

function handleFile(file) {
  clearError();
  if (!file) return;
  SOURCE = file.name;
  const str = S();
  showLoading(str.loadReading);              // hide uploader, show the bar immediately

  const reader = new FileReader();
  // read phase drives the first stretch of the bar (0–35%)
  reader.onprogress = e => { if (e.lengthComputable) setProgress(5 + (e.loaded / e.total) * 30, str.loadReading); };
  reader.onerror = () => { MODEL = null; render(); showError(S().errParse); };
  reader.onload = async ev => {
    try {
      // each blocking step: bump the bar + label, let it PAINT, then do the work
      setProgress(40, str.loadParsing); await yieldPaint();
      const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: false });
      setProgress(70, str.loadComputing); await yieldPaint();
      const data = cleanData(wb);
      const model = computeModel(data);
      MODEL = model; SOURCE = file.name;
      setProgress(100, str.loadRendering); await yieldPaint();
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } catch (err) {
      console.error(err);
      const s = S();
      MODEL = null; render();                // restores uploader, hides the bar
      if (err && err.key === 'errNoSheet') showError(fmt(s.errNoSheet, { sheet: err.sheet }));
      else if (err && err.key === 'errNoSales') showError(s.errNoSales);
      else showError(s.errParse);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ------------------------------------------------------------------ *
 * 7. Wiring
 * ------------------------------------------------------------------ */
function init() {
  document.documentElement.setAttribute('data-theme', THEME);
  document.documentElement.setAttribute('lang', LANG);

  const fileInput = $('file-input'), drop = $('drop');
  $('btn-browse').addEventListener('click', () => fileInput.click());
  drop.addEventListener('click', e => { if (e.target.closest('#btn-browse')) return; fileInput.click(); });
  fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); if (ev === 'dragleave' && drop.contains(e.relatedTarget)) return; drop.classList.remove('drag'); }));
  drop.addEventListener('drop', e => { const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handleFile(f); });

  $('btn-reupload').addEventListener('click', () => { MODEL = null; SOURCE = ''; $('file-input').value = ''; if (location.hash) history.replaceState(null, '', location.pathname + location.search); render(); });
  $('btn-share').addEventListener('click', shareCurrent);
  $('btn-sample').addEventListener('click', downloadSample);

  document.querySelectorAll('#lang-seg .seg-btn').forEach(b => b.addEventListener('click', () => { LANG = b.dataset.lang; localStorage.setItem('mro.lang', LANG); render(); }));
  document.querySelectorAll('#theme-seg .seg-btn').forEach(b => b.addEventListener('click', () => { THEME = b.dataset.themeBtn; localStorage.setItem('mro.theme', THEME); render(); }));

  window.addEventListener('resize', updateScrollShadows);

  // shared link? decode and render from the URL; otherwise show the uploader
  if (/^#d=/.test(location.hash)) loadShared(location.hash.slice(3));
  else render();
}
document.addEventListener('DOMContentLoaded', init);
