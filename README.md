# Maxima Executive Dashboard

A static, client-side executive dashboard for **MRO International**. Upload the
periodic data workbook (`.xlsx`) and the page parses it **in the browser** and
renders the executive summary — sales vs. last year, markets & products, and the
open-order backlog (backorders and awaiting-fulfillment). No server, no upload:
every figure is computed deterministically from the file you choose.

Built to run on **GitHub Pages** (or any static host).

## Features

- **Upload-driven** — drag & drop or pick an `.xlsx`; nothing is sent anywhere.
- **Deterministic** — all KPIs, charts, and breakdowns are computed from the raw
  sheets on each load. Nothing is hard-coded.
- **Bilingual** — Español / English toggle (top right). Default: Español.
- **Light / dark** — theme toggle (top right). Default: light ("clear").
- **Self-contained** — the XLSX parser ([SheetJS](https://sheetjs.com)) is
  vendored under `assets/vendor/`, so the app works offline once loaded.
- Language and theme choices persist across visits (via `localStorage`).

## Expected workbook format

The uploaded `.xlsx` must contain these sheets (column headers matched by name,
so column order can vary):

| Sheet | Purpose | Key columns |
|-------|---------|-------------|
| `Sales` | Invoiced sales lines | `ID`, `Customer`, `Sales Rep`, `Transaction Type`, `Date`, `Document Number`, `Item`, `Amount`, `Customer/Project: Customer Type`, `Inventory Location: Name`, `Address: Billing Address Country Name`, `Product Segment`, `SO Date` |
| `Target` | Sales goals by customer | `ID`, `Amount`, `Country`, `Customer Type`, `Type` |
| `PendingFullFill` | Open-order backlog | `Item`, `Date`, `Document Number`, `Customer`, `Ordered`, `Fulfilled`, `Committed`, `Back Ordered`, `Aggregate Amount`, `Unit Price`, `Location`, `Validated Status` |
| `IF` | Item fulfillments | `Internal ID`, `Subsidiary`, `Date`, `Document Number`, `Status`, `Created` |

`Sales` and `PendingFullFill` are **required**; `Target` and `IF` are parsed if
present (`Target` is parsed but not yet displayed). The sheet named `Target` here
was `26_Target` in the original source report.

## How the numbers are computed

The build applies the same cleaning rules as the source report, then computes:

- **Cleaning** — `Sales` uses the `Amount` column only; rows for Sales Rep
  *Oscar Riquelme*, account *CMM02359*, and `Subtotal` item lines are excluded;
  warehouse codes `VTNWH-MX-EU`/`VTNWH-MX` are unified to `VTNWH-MX`; missing
  `Product Segment` is back-filled from the most common segment per item/memo.
- **As-of date** — the latest date found in `Sales`. Year-to-date runs from
  Jan 1 of that year through the as-of date; "same period last year" is the same
  window shifted back one year. The monthly bar chart shows complete months only
  (the partial current month is excluded but still counted in the YTD total).
- **Sales KPIs** — YTD net sales, same-period prior year, Δ %, active customers,
  and country count.
- **Markets** — top countries and top product segments by YTD `Amount`.
- **Open order book** (all `PendingFullFill` rows) —
  open value `Σ Aggregate Amount`; committed/ready value `Σ (Unit Price × Committed)`;
  backorder value `Σ (Unit Price × Back Ordered)`; distinct SOs and customers;
  breakdowns by warehouse and by customer.

## Run locally

```bash
# from the repo root
python3 -m http.server 8080
# open http://localhost:8080
```

(Opening `index.html` directly via `file://` also works in most browsers, but a
local server matches the GitHub Pages environment.)

## Deploy to GitHub Pages

1. Push this repo to `github.com/MatiMaster/MaximaExecutiveDashboard`.
2. In **Settings → Pages**, set the source to the `main` branch, root folder.
3. The site publishes at `https://matimaster.github.io/MaximaExecutiveDashboard/`.

The `.nojekyll` file ensures the `assets/` folder is served as-is.

## Project structure

```
index.html            # page shell + upload UI
assets/
  styles.css          # light (default) & dark themes
  i18n.js             # ES/EN strings + label dictionaries
  app.js              # parse → compute → render
  vendor/
    xlsx.full.min.js  # SheetJS (vendored)
```
