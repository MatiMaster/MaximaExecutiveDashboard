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

const $ = id => document.getElementById(id);
const S = () => I18N.STR[LANG];
const fmt = (tpl, vars) => String(tpl).replace(/\{(\w+)\}/g, (_, k) => (vars && k in vars) ? vars[k] : '');

/* ------------------------------------------------------------------ *
 * 1. Formatting (locale-aware)
 * ------------------------------------------------------------------ */
const locale = () => LANG === 'es' ? 'es-CL' : 'en-US';
const full = n => '$' + Math.round(Number(n) || 0).toLocaleString(locale());
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

  // language / theme button states
  document.querySelectorAll('#lang-seg .seg-btn').forEach(b => b.classList.toggle('active', b.dataset.lang === LANG));
  document.querySelectorAll('#theme-seg .seg-btn').forEach(b => b.classList.toggle('active', b.dataset.themeBtn === THEME));

  if (!MODEL) { renderEmpty(); return; }

  const m = MODEL, sd = m.sales, bk = m.book, mk = m.markets;

  // as-of box
  $('asof-box').hidden = false;
  $('asof-big').textContent = fmt(str.asOfBig, { date: longDate(m.asOf) });
  $('asof-mut').textContent = fmt(str.asOfMut, { start: shortDM(m.yearStart), date: shortDM(m.asOf) });
  $('btn-reupload').hidden = false;

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
  </div>`;

  // charts + composition bar (after DOM insert; they read theme CSS vars)
  $('chart-trend').innerHTML = monthlyBarsSVG(sd.monthly);
  $('chart-cum').innerHTML = cumulativeSVG(sd.monthly);
  const tot = bk.readyValue + bk.boValue || 1, rp = bk.readyValue / tot * 100, bp = bk.boValue / tot * 100;
  $('comp').innerHTML =
    `<span style="flex:${rp};background:var(--ready);color:var(--on-neutral)">${Math.round(rp)}%</span>` +
    `<span style="flex:${bp};background:var(--bo);color:#fff">${Math.round(bp)}%</span>`;
  wireTips('#chart-trend rect'); wireTips('#chart-cum circle');

  // footer
  $('foot').hidden = false;
  $('foot').innerHTML = fmt(str.foot, { source: esc(SOURCE || '—'), date: longDate(m.asOf) });
}

function renderEmpty() {
  $('asof-box').hidden = true;
  $('btn-reupload').hidden = true;
  $('dashboard').hidden = true;
  $('foot').hidden = true;
  $('uploader').hidden = false;
  const str = S();
  $('up-title').textContent = str.upTitle;
  $('up-sub').innerHTML = str.upSub;
  $('btn-browse').textContent = str.upBrowse;
  $('up-fmt').innerHTML = str.upFmt;
}

/* ------------------------------------------------------------------ *
 * 6. File handling
 * ------------------------------------------------------------------ */
function showError(msg) { const e = $('up-err'); e.hidden = false; e.textContent = msg; }
function clearError() { $('up-err').hidden = true; }

function handleFile(file) {
  clearError();
  if (!file) return;
  SOURCE = file.name;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: false });
      const data = cleanData(wb);
      MODEL = computeModel(data);
      $('uploader').hidden = true;
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } catch (err) {
      console.error(err);
      const str = S();
      if (err && err.key === 'errNoSheet') showError(fmt(str.errNoSheet, { sheet: err.sheet }));
      else if (err && err.key === 'errNoSales') showError(str.errNoSales);
      else showError(str.errParse);
    }
  };
  reader.onerror = () => showError(S().errParse);
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

  $('btn-reupload').addEventListener('click', () => { MODEL = null; SOURCE = ''; $('file-input').value = ''; render(); });

  document.querySelectorAll('#lang-seg .seg-btn').forEach(b => b.addEventListener('click', () => { LANG = b.dataset.lang; localStorage.setItem('mro.lang', LANG); render(); }));
  document.querySelectorAll('#theme-seg .seg-btn').forEach(b => b.addEventListener('click', () => { THEME = b.dataset.themeBtn; localStorage.setItem('mro.theme', THEME); render(); }));

  render();
}
document.addEventListener('DOMContentLoaded', init);
