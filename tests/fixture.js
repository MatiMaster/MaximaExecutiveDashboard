/* ============================================================
 * Hand-crafted fixture: a tiny dataset where EVERY displayed value can
 * be computed by hand (see EXPECTED, with the arithmetic shown). This is
 * the anti-regression anchor — if a calculation changes, these break.
 *
 * As-of date = latest Sales date = 2026-07-01  ->  year 2026, prev 2025.
 * YTD window        = [2026-01-01, 2026-07-01]
 * Same-period LY    = [2025-01-01, 2025-07-01]
 * Full prior year   = [2025-01-01, 2025-12-31]
 * ============================================================ */
'use strict';

const SALES_HEAD = ['ID', 'Customer', 'Sales Rep', 'Transaction Type', 'Date', 'Document Number', 'Item', 'Memo', 'Subsidiary: Name', 'Customer Category: Name', 'Quantity', 'Unit Price', 'Amount (Foreign Currency)', 'Amount', 'Customer/Project: Customer Type', 'Inventory Location: Name', 'Address: Billing Address Country Name', 'Product Segment', 'SO Date'];

// Amount (Foreign Currency) is deliberately a large decoy (999) on every row —
// if the parser ever used it instead of Amount, YTD would explode.
const SALES = [SALES_HEAD,
  // in-window 2026 (YTD)
  ['C1', 'Cust A', 'Rep One', 'Invoice', '2026-02-10', 'INV1', 'SKU1', 'memo1', 'Sub', 'CatX', 10, 5, 999, 100, 'EX-ASIA', 'MXWHS', 'Chile', '4T Engine Oils', ''],
  ['C1', 'Cust A', 'Rep One', 'Invoice', '2026-03-15', 'INV2', 'SKU2', 'memo2', 'Sub', 'CatX', 20, 10, 999, 200, 'EX-ASIA', 'MXWHS', 'Chile', '2T Engine Oils', ''],
  ['C2', 'Cust B', 'Rep Two', 'Invoice', '2026-06-20', 'INV3', 'SKU1', 'memo1', 'Sub', 'CatY', 5, 5, 999, 300, 'EX-LATAM', 'EU Warehouse', 'Germany', '4T Engine Oils', ''],
  ['C2', 'Cust B', 'Rep Two', 'Invoice', '2026-07-01', 'INV4', 'SKU3', 'memo3', 'Sub', 'CatY', 4, 25, 999, 50, 'EX-LATAM', 'EU Warehouse', 'Germany', 'Aerosols', ''], // July = partial current month (in YTD, out of monthly bars)
  ['C4', 'Cust D', 'Rep Two', 'Invoice', '2026-05-05', 'INV10', 'SKU2', 'memo2', 'Sub', 'CatX', 2, 10, 999, 400, 'EX-ASIA', 'VTNWH-MX-EU', 'Japan', '2T Engine Oils', ''], // loc must normalize to VTNWH-MX
  // 2025 rows
  ['C1', 'Cust A', 'Rep One', 'Invoice', '2025-02-10', 'INV5', 'SKU1', 'memo1', 'Sub', 'CatX', 10, 5, 999, 80, 'EX-ASIA', 'MXWHS', 'Chile', '4T Engine Oils', ''], // same-period LY + full prior year
  ['C1', 'Cust A', 'Rep One', 'Invoice', '2025-11-10', 'INV6', 'SKU1', 'memo1', 'Sub', 'CatX', 10, 5, 999, 120, 'EX-ASIA', 'MXWHS', 'Chile', '4T Engine Oils', ''], // full prior year only (Nov > Jul)
  // rows that MUST be excluded by the cleaning rules
  ['C3', 'Cust C', 'Oscar Riquelme', 'Invoice', '2026-04-10', 'INV7', 'SKU1', 'memo1', 'Sub', 'CatZ', 100, 5, 999, 9999, 'EX-EU', 'MXWHS', 'France', '4T Engine Oils', ''], // excluded rep
  ['CMM02359', 'CMM02359 ACCOUNT', 'Rep One', 'Invoice', '2026-04-11', 'INV8', 'SKU1', 'memo1', 'Sub', 'CatX', 1, 1, 999, 8888, 'EX-ASIA', 'MXWHS', 'Chile', '4T Engine Oils', ''], // excluded account
  ['C1', 'Cust A', 'Rep One', 'Invoice', '2026-04-12', 'INV9', 'Subtotal', '', 'Sub', 'CatX', 0, 0, 999, 7777, 'EX-ASIA', 'MXWHS', 'Chile', '', ''] // excluded subtotal line
];

const TARGET = [['ID', 'Amount', 'Country', 'Customer Type', 'Type'],
  ['C1', 50000, 'Chile', 'EX-ASIA', 'Target'],
  ['C2', 30000, 'Germany', 'EX-LATAM', 'Target']];

// PendingFullFill has NO Product Segment column here on purpose, to exercise
// the segment back-fill from the Sales item->segment map.
const PFF_HEAD = ['Item Type', 'Item', 'Description (Sales)', 'Date', 'Transaction Type', 'Primary Sales Rep', 'Document Number', 'Customer', 'Ordered', 'Fulfilled', 'Committed', 'Back Ordered', 'Subsidiary: Name', 'Aggregate Amount', 'Unit Price', 'Location', 'Validated Status'];
const PFF = [PFF_HEAD,
  ['Assembly', 'SKU1', 'memo1', '2026-05-01', 'Sales Order', 'Rep One', 'SO1', 'Cust A', 100, 0, 60, 40, 'Sub', 500, 5, 'MXWHS', 'Pending Fulfillment'],
  ['Assembly', 'SKU2', 'memo2', '2026-05-02', 'Sales Order', 'Rep Two', 'SO2', 'Cust B', 50, 0, 50, 0, 'Sub', 500, 10, 'EU Warehouse', 'Partially Fulfilled'],
  ['Assembly', 'SKU1', 'memo1', '2026-05-03', 'Sales Order', 'Rep One', 'SO1', 'Cust A', 20, 0, 10, 10, 'Sub', 100, 5, 'VTNWH-MX-EU', 'Pending Fulfillment'], // same SO1 doc; loc normalizes
  ['Assembly', 'SKU3', 'memo3', '2026-05-04', 'Sales Order', 'Rep Two', 'SO3', 'Cust D', 8, 0, 0, 8, 'Sub', 200, 25, 'MXWHS', 'Pending Fulfillment']];

const IF = [['Internal ID', 'Subsidiary', 'Date', 'Type', 'Document Number', 'Transaction Number', 'Name', 'PO/Check Number', 'Status', 'Memo', 'Amount (Foreign Currency)', 'Amount', 'Category', 'Sales Rep', 'Location', 'Shipping Country', 'Terms', 'Created'],
  [1, 'MRO Holding : Sub', '2026-06-01', 'Item Fulfillment', 'IF1', 'TS1', 'Cust A', '', 'Released', '', 0, 0, 'Retail', 'Rep One', 'MXWHS', 'Chile', '', 'SO1']];

const SHEETS = { Sales: SALES, Target: TARGET, PendingFullFill: PFF, IF };

// Every number below is derived by hand from the rows above.
const EXPECTED = {
  cleaning: { salesRows: 7 },                       // 10 rows - Oscar - CMM02359 - Subtotal
  asOf: '2026-07-01', year: 2026, prev: 2025,
  sales: {
    ytd: 1050,        // 100+200+300+50+400
    ly: 80,           // only INV5 (2025-02); INV6 is Nov, out of same-period
    delta: 970,       // 1050-80
    deltaPct: 970 / 80,
    activeCustomers: 3, // A,B,D
    countries: 3,       // Chile,Germany,Japan
    monthlyIdx: [1, 2, 3, 4, 5, 6],
    monthlyCur: [0, 100, 200, 0, 400, 300], // Jul (50) excluded as partial month
    monthlyPrev: [0, 80, 0, 0, 0, 0],
    monthlyCurSum: 1000,                    // != ytd(1050): the extra 50 is July
    partialMonth: 7
  },
  markets: {
    topCountries: [['Japan', 400], ['Germany', 350], ['Chile', 300]],
    topSegments: [['2T Engine Oils', 600], ['4T Engine Oils', 400], ['Aerosols', 50]],
    topCountriesShare: 1050 / 1050, // all 3 countries fit in top-7
    topSegShare: 600 / 1050
  },
  book: {
    openValue: 1300,      // Σ agg = 500+500+100+200
    openOrders: 3,        // distinct docs SO1,SO2,SO3 (SO1 appears twice)
    bookCustomers: 3,
    readyValue: 850,      // Σ up*com = 300+500+50+0
    readyUnits: 120,      // Σ com
    orderedUnits: 178,    // Σ ord
    boValue: 450,         // Σ up*bo (bo>0) = 200+50+200
    boUnits: 58,          // 40+10+8
    boSOs: 2,             // distinct docs with a backorder: SO1, SO3
    readyShare: 850 / 1300,
    status: [['Pending Fulfillment', 2], ['Partially Fulfilled', 1]],
    boWh: [['MXWHS', 400], ['VTNWH-MX', 50]],
    boCust: [['Cust A', 250], ['Cust D', 200]],
    readyWh: [['EU Warehouse', 500], ['MXWHS', 300], ['VTNWH-MX', 50]],
    readyCust: [['Cust B', 500], ['Cust A', 350], ['Cust D', 0]]
  },
  // per-item and per-segment (ytd, units, ly, fy, bo)
  prodPerf: {
    SKU1: { ytd: 400, units: 15, ly: 80, fy: 200, bo: 250, boUnits: 50, seg: '4T Engine Oils', name: 'memo1' },
    SKU2: { ytd: 600, units: 22, ly: 0, fy: 0, bo: 0, boUnits: 0, seg: '2T Engine Oils', name: 'memo2' },
    SKU3: { ytd: 50, units: 4, ly: 0, fy: 0, bo: 200, boUnits: 8, seg: 'Aerosols', name: 'memo3' }
  },
  segPerf: {
    '4T Engine Oils': { ytd: 400, units: 15, ly: 80, fy: 200, bo: 250 },
    '2T Engine Oils': { ytd: 600, units: 22, ly: 0, fy: 0, bo: 0 },
    'Aerosols': { ytd: 50, units: 4, ly: 0, fy: 0, bo: 200 }
  }
};

module.exports = { SHEETS, EXPECTED };
