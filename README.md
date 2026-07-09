# Maxima Executive Dashboard

A **static, client-side executive dashboard** for **MRO International** (the
international division of Maxima Racing Oils). A user uploads the periodic
NetSuite export (`.xlsx`); the page parses it **entirely in the browser** and
renders an executive summary — sales vs. last year, markets & products, the
open-order backlog, and ranked product performance.

**No backend.** All parsing, calculation, and rendering happen in the browser.
It's designed to be served from **GitHub Pages** (or any static host) at
`https://matimaster.github.io/MaximaExecutiveDashboard/`.

---

## For anyone (or any LLM) picking this up

This section is the fast path to understanding the project.

### What it is and why it exists
- It replaces a **static, hand-numbered CEO snapshot** (`mro_ceo_snapshot.html`
  in the parent workspace) whose figures were baked in. This project computes
  those same figures **dynamically from an uploaded file**, so the dashboard is
  always current and never hand-edited.
- The calculation rules were reverse-engineered from a detailed internal tool
  (`mro_int_dashboard_*.html`) so the executive view matches the source of truth.
- Audience: a Spanish-speaking executive. Spanish is the default language.

### Architecture / data flow
```
upload .xlsx ─▶ FileReader ─▶ XLSX.read (SheetJS)
             ─▶ cleanData()     apply source rules, normalize, back-fill segments
             ─▶ computeModel()  produce every displayed number (the "model")
             ─▶ render()        build the HTML/SVG from the model
```
- `cleanData(workbook)` → `{ sales, pff, target, ifs }` of normalized row objects.
- `computeModel(data)` → the **model**: a plain object holding every value the UI
  shows (KPIs, monthly series, markets, the order book, per-product/segment
  tables). It is **pure and deterministic** — the "as-of" date is the latest
  date found in `Sales`, never `Date.now()`, so the same file always yields the
  same numbers. **This is the function tests target.**
- `render()` reads the model + current language/theme and writes the DOM. It
  computes only display-derived things (deltas, percentages, formatting).

### Files
```
index.html              page shell, header, upload UI, loading bar
assets/
  app.js                parse → clean → computeModel → render (all logic)
  i18n.js               ES/EN strings + label dictionaries (segment/warehouse/…)
  styles.css            light (default) + dark themes
  favicon.png           Maxima tab icon (fetched from maximaeurope.eu)
  MaximaRacingOilsLogo2.png   header wordmark (transparent background)
  vendor/xlsx.full.min.js     SheetJS, vendored (works offline)
tests/                  Node test suite — see "Testing" below
package.json            only exists to run the tests (`npm test`)
```

### Features
- **Upload-driven, deterministic** — every figure is computed from the file; nothing stored.
- **Filters** (live upload only) — a filter bar slices the whole dashboard: Sales Rep, Client Category (International/Domestic, derived from the rep), Subsidiary, Location/warehouse, Customer, Customer Type, Customer Category, Country, Product Segment (multi-select each), a Document-number exact match, and a Date From/To range. Filtering re-derives the model from the cleaned data via `filterData()` → `computeModel(fd, {asOf, from, to})`, so **every** section (sales KPIs, markets, order book, funnel, backorders, product/segment performance, classification) moves together. Defaults are empty ⇒ identical to the full dataset. The as-of anchor is pinned to the full-data max Sales date so filtering never shifts the period; the date range, when set, redefines the current period and shifts the same-period-LY window with it, and also clips the open book by SO date. Filters live on the raw rows, so they are **not** available on shared links / snapshots (which carry only the computed model — the bar hides itself there).
- **Bilingual** — Español (default) / English toggle. Numbers/percentages/dates are locale-formatted; segment, warehouse, status, and common country labels translate via `i18n.js` dictionaries (unknown values pass through).
- **Light (default) / dark** themes.
- **Shareable links** — the **Share** button embeds the *computed model* (not the raw rows) in the URL hash: compact-encode → gzip (`CompressionStream`) → base64url → `#d=…`. Opening the link decodes and renders with no backend; the hash is never sent to the server, so data stays client-side. ~31 KB for a full real dataset.
- **Product classification** (Section 7) — buckets the top products into doing-great / OK / slow-declining by growth vs. last year, with count cards and per-bucket lists; has its own backorder toggle.
- **Sample file** — "Download a sample file" generates a valid `.xlsx` (obfuscated dummy data, realistic magnitudes) in the browser so users see the exact expected format.
- **Progress bar** on upload — parsing is synchronous and briefly blocks; the bar advances through Reading → Parsing → Computing → Building so the page never looks frozen.

---

## Expected workbook format

Sheets (columns matched by header name, so column order can vary):

| Sheet | Required | Purpose | Key columns |
|-------|----------|---------|-------------|
| `Sales` | ✅ | Invoiced sales lines | `ID`, `Customer`, `Sales Rep`, `Transaction Type`, `Date`, `Document Number`, `Item`, `Memo`, `Amount`, `Customer/Project: Customer Type`, `Inventory Location: Name`, `Address: Billing Address Country Name`, `Product Segment`, `SO Date` |
| `PendingFullFill` | ✅ | Open-order backlog | `Item`, `Date`, `Document Number`, `Customer`, `Ordered`, `Fulfilled`, `Committed`, `Back Ordered`, `Aggregate Amount`, `Unit Price`, `Location`, `Validated Status` |
| `Target` | optional | Sales goals by customer id (parsed, not yet displayed) | `ID`, `Amount`, `Country`, `Customer Type`, `Type` |
| `IF` | optional | Item fulfillments (parsed, not displayed) | `Internal ID`, `Date`, `Document Number`, `Status`, `Created` |

The sheet named `Target` was `26_Target` in the original source report (both names are accepted).

---

## What every displayed number means (calculation reference)

The "as-of" date is the latest `Sales` date; `year` = its year, `prev` = year−1.
Windows: **YTD** = `[year-01-01, asOf]`; **same-period LY** = the YTD window
shifted back one year; **full prior year** = `[prev-01-01, prev-12-31]`.

**Cleaning (applied before anything else):** use the `Amount` column (never
`Amount (Foreign Currency)`); exclude rows for Sales Rep *Oscar Riquelme*,
account *CMM02359*, and `Item = Subtotal`; normalize `VTNWH-MX-EU → VTNWH-MX`;
back-fill a missing `Product Segment` from the most common segment per item/memo.

| Section | Value | Definition |
|---|---|---|
| 1 Sales | YTD net sales | Σ `Amount` in YTD window |
| | Same period LY | Σ `Amount` in same-period-LY window |
| | Δ / Δ% | `ytd − ly` and `(ytd − ly) / ly` |
| | Active customers / countries | distinct `Customer` / `Country` in YTD |
| | Monthly bars | Σ `Amount` per **complete** month of the year; the partial current month is excluded from the bars but still counted in YTD |
| | Cumulative | running sum of the monthly series |
| 2 Markets | Top countries / segments | Σ `Amount` YTD grouped, sorted desc (top 7 / top 6) |
| | Shares | top-countries sum ÷ YTD; top segment ÷ YTD |
| 3 Open book | Open value | Σ `Aggregate Amount` (all PFF rows) |
| | Open orders / customers | distinct `Document Number` / `Customer` |
| | Committed (ready) value / units | Σ `Unit Price × Committed` / Σ `Committed` |
| | Backorder value | Σ `Unit Price × Back Ordered` where `Back Ordered > 0` |
| | Status chips | distinct SO count per `Validated Status`; plus Σ `Ordered` |
| 3 Funnel | Open SO | qty = distinct pending `Document Number`; amount = Σ `Aggregate Amount`; avg days = mean (as-of − row date) |
| | Waiting IF / IF In Progress | split each pending SO by whether the **IF sheet** has a row whose `Created` = that SO. Waiting = no IF; In Progress = has IF. `IF Qty` = distinct IF docs whose source SO is still pending; `SO Qty` = distinct pending SOs with an IF; amount = Σ `Aggregate Amount` of those SOs; avg days over IF dates. Open SO amount = Waiting + In Progress (a partition) |
| | BO | units = Σ `Back Ordered`; amount = Σ `Unit Price × Back Ordered`; % of open value; avg days over backordered lines |
| 4 Backorders | Value / units / SOs | as above; units = Σ `Back Ordered`; SOs = distinct docs with a backorder |
| | By warehouse / customer | grouped Σ `Unit Price × Back Ordered`, **top 5** |
| 5 Awaiting fulfillment | Ready value / units / % | committed value/units; % = ready ÷ open value |
| | By warehouse / customer | grouped Σ `Unit Price × Committed`, **top 5** |
| 6 Product performance | Per product & per segment | YTD sales, units (Σ `Quantity` YTD), same-period-LY, full-prior-year |
| | Backorder-adjusted (toggle) | `Adjusted = YTD sales + Σ (Unit Price × Back Ordered)`, re-ranked; per item/segment |
| 7 Classification | Growth % | `(sales − LY) / LY` per product (LY = same-period last year); `sales` = YTD, or the adjusted total when the section's backorder toggle is on; no prior-year sales ⇒ 0% |
| | Buckets | over the **top 100 by sales** (segment required): doing-great ≥ +15%, OK strictly between, slow/declining ≤ −15% |

Products are keyed by **item SKU** (shared between `Sales` and `PendingFullFill`);
the displayed name is that SKU's most common memo/description.

---

## Testing — REQUIRED for all calculations

> **Accuracy of the displayed values is the top priority of this project.**
> Every calculation that reaches the screen must be covered by a test, and
> **any change to a calculation (or any new displayed value) must ship with a
> test in the same change.** Do not merge calculation changes without tests.

### Run
```bash
npm test          # or: node --test
```
Requires Node 18+ (uses the built-in test runner, `node:assert`, and the web
`CompressionStream`/`Blob` globals). No dependencies to install.

### How it works
`tests/helper.js` loads the real browser modules (`i18n.js`, vendored SheetJS,
`app.js`) into a Node VM with lightweight DOM stubs and exposes the pure
functions (`cleanData`, `computeModel`, formatters, share codec, sample builder)
— so tests exercise **the exact code the page runs**, not a reimplementation.

### Coverage (62 tests)
- **`tests/fixture.js` + `tests/calculations.test.js`** — a tiny hand-crafted
  dataset where **every displayed number is computed by hand** (the arithmetic
  is shown in `EXPECTED`), asserted section by section. This is the anti-
  regression anchor.
- **`tests/cleaning.test.js`** — the exclusion rules, Amount-vs-foreign-currency,
  `VTNWH-MX` normalization, segment back-fill, date parsing (ISO / Excel serial),
  and the same-period-last-year shift.
- **`tests/invariants.test.js`** — structural checks on the sample data with **no
  hardcoded numbers** (per-product/segment YTD sums to total; backorder value is
  fully attributed across products and segments; ratios equal their definitions;
  ranked lists sorted; series lengths aligned). These catch double-counting,
  dropped rows, and mis-attribution generically.
- **`tests/filters.test.js`** — every filter dimension narrows the model correctly
  (sales KPIs, order book, and funnel all respond), client-category-from-rep,
  the date range redefining the period + shifting LY + clipping the book, filters
  composing, `filterOptions` menus, and "no filters ⇒ full-data model exactly".
- **`tests/funnel.test.js`** — the Section 3 fulfillment funnel: the PFF↔IF join
  by the IF's `Created` (source SO), each card (Open SO / Waiting IF / IF In
  Progress / BO), avg-days over the right date set, and the Open = Waiting +
  In-Progress partition.
- **`tests/classification.test.js`** — the Section 7 buckets (±15% thresholds
  incl. boundaries, no-prior-year ⇒ OK, segment/activity exclusions, sales-desc
  ranking, and the backorder toggle moving a product between buckets).
- **`tests/formatting.test.js`** — the locale formatters in ES and EN (grouping,
  decimals, M/K abbreviation, signs, percentages) — what the executive reads.
- **`tests/sharelink.test.js`** — the share codec round-trips the model
  value-for-value (compact→hydrate and gzip encode→decode), so shared links can
  never silently show different figures.

### The rule for future work
When you add or change a displayed metric:
1. Add its exact definition to the calculation reference table above.
2. Add hand-computed expected values to `tests/fixture.js` (extend the fixture
   rows if needed) and assert them in `tests/calculations.test.js`.
3. If it's an aggregate, add a structural invariant in `tests/invariants.test.js`
   (e.g. it sums/relates to a total).
4. Run `npm test` — it must be green before shipping.

---

## Run locally
```bash
python3 -m http.server 8080   # then open http://localhost:8080
```

## Deploy (GitHub Pages)
Push to `main`; in **Settings → Pages** set source = `main` / root. `.nojekyll`
serves `assets/` as-is. After a deploy, browsers may cache old assets — hard-
refresh (Cmd/Ctrl+Shift+R) to see changes immediately.

## Developer notes / gotchas
- **Determinism:** `computeModel` derives the as-of date from the data, not the
  clock — keep it that way so tests stay stable.
- **Background tabs:** `requestAnimationFrame` is paused in unfocused tabs, so
  the loading yield races rAF against a timeout fallback (don't remove it, or
  uploads can hang for a backgrounded tab).
- **Header width:** the one-row header is width-tuned; the wide wordmark logo
  only fits because the as-of line is stacked under the brand text (not in the
  right-side controls). The share button has a fixed `min-width` so its label
  changes never reflow the header.
- **Tests & realms:** objects from the VM realm fail `deepStrictEqual` on
  prototype identity — normalize with `JSON.parse(JSON.stringify(...))` before
  deep-comparing (see `calculations.test.js`).
