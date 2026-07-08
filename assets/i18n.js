/* ============================================================
   i18n — UI strings (ES/EN) + deterministic label dictionaries.
   Data values are shown as-is; only the enum-like fields
   (segment, warehouse, validated status, common countries) are
   translated through a fixed dictionary, falling back to raw.
   ============================================================ */
(function (root) {
  'use strict';

  // ---- UI strings. {tokens} are filled by fmt() in app.js ----
  const STR = {
    es: {
      title: 'MRO International — Resumen Ejecutivo',
      subtitle: 'Ventas internacionales y cartera de pedidos · monto neto facturado',
      asOfBig: 'Al {date}',
      asOfMut: 'Acumulado del año: {start} – {date}',
      reupload: 'Cambiar archivo',
      share: 'Compartir', shareBuilding: 'Generando…', shareCopied: '✓ Enlace copiado', shareErr: 'Error al generar',
      loadingShared: 'Cargando datos compartidos…',
      errShared: 'El enlace compartido no es válido o está dañado.',

      upTitle: 'Cargar archivo de datos',
      upSub: 'Arrastra un archivo <b>.xlsx</b> aquí o haz clic para seleccionarlo.',
      upBrowse: 'Seleccionar archivo',
      upFmt: 'El archivo debe contener las hojas <code>Sales</code>, <code>Target</code>, <code>PendingFullFill</code> e <code>IF</code>, con el mismo formato de columnas del reporte original. Todo el procesamiento ocurre en tu navegador — nada se sube a internet.',
      errNoSheet: 'No se encontró la hoja requerida: {sheet}.',
      errParse: 'No se pudo leer el archivo. Verifica que sea un .xlsx válido con el formato esperado.',
      errNoSales: 'La hoja Sales no contiene filas de ventas válidas.',

      s1h: 'Ventas — este año vs. año anterior',
      s1st: 'Ventas netas acumuladas, comparadas con el mismo período de {prev}',
      s1curLabel: 'Ventas netas {year} · acumulado del año',
      s1curNote: '{start} – {date}, {year} · <b>{full}</b>',
      s1prevLabel: 'Mismo período {prev}',
      s1prevNote: '{start} – {date}, {prev}',
      s1custLabel: 'Clientes activos',
      s1custNote: 'en <b>{n}</b> países · acum. año',
      pillVs: '{arrow} {sign}{pct}  ({dsign}{delta} vs. {prev})',
      legMonthly: 'Ventas netas mensuales · {month} excluido (mes parcial)',
      legCumNote: 'Ventas netas acumuladas en el año · ¿vamos por delante o por detrás?',
      legCumCur: '{year} acumulado',
      legCumPrev: '{prev} acumulado',

      s2h: 'Mercados y productos',
      s2st: 'De dónde provienen las ventas · acumulado del año {year}',
      s2ctryLabel: 'Ventas por país',
      s2ctryNote: 'Top mercados internacionales · ≈ {pct} de las ventas',
      s2segLabel: 'Mezcla de productos',
      s2segNote: 'Ventas por línea · {top} ≈ {pct} del total',

      s3h: 'Pedidos de venta abiertos',
      s3st: 'Pedidos registrados y aún no despachados por completo',
      s3valLabel: 'Valor de pedidos abiertos',
      s3valNote: '<b>{n}</b> pedidos abiertos · <b>{m}</b> clientes',
      s3compLabel: 'Composición de la cartera de pedidos',
      s3ready: 'Por despachar',
      s3bo: 'En backorder',
      s3units: 'Unidades pedidas',

      s4h: 'Backorders',
      s4st: 'Pedido pero a la espera de stock',
      s4valLabel: 'Valor en backorder · a la espera de stock',
      s4units: 'Unidades',
      s4sos: 'Pedidos afectados',

      s5h: 'En espera de despacho',
      s5st: 'Comprometido y asignado — en espera de Item Fulfillment (IF)',
      s5valLabel: 'Listo para despachar · valor comprometido',
      s5unitsLabel: 'Unidades asignadas',
      s5shareLabel: '% del valor abierto',

      byWh: 'Por bodega',
      byCust: 'Mayores por cliente',

      s6h: 'Rendimiento de productos',
      s6st: 'Cómo se venden los productos · ranking por ventas del año — activa el backorder para ver el impacto',
      perfView: 'Vista', perfActual: 'Ventas actuales', perfBO: 'Con backorder',
      tblSeg: 'Comparación por segmento',
      tblProd: 'Comparación por producto',
      abbrPerf: '<b>MP</b> = mismo período {prev} · <b>AC</b> = año completo {prev} · <b>Ventas año</b> = acumulado {year}',
      abbrPerfBO: '<b>Ajustado</b> = Ventas año + Backorder (precio unit. × unidades en backorder) · re-ordenado por Ajustado',
      colProduct: 'Producto', colSegment: 'Segmento', colYTD: 'Ventas año', colUnits: 'Unid.',
      colSP: 'MP {prev}', colDSP: 'Δ MP', colPSP: '% MP', colFY: 'AC {prev}', colDFY: 'Δ AC', colPFY: '% AC',
      colBO: 'Backorder', colAdj: 'Ajustado',
      tipRank: 'Posición en el ranking',
      tipProduct: 'Código (SKU) y nombre del producto',
      tipSegment: 'Línea o segmento de producto',
      tipYTD: 'Ventas año — ventas netas acumuladas del año, del 1 de enero a la fecha de corte',
      tipUnits: 'Unidades vendidas en el período',
      tipBO: 'Backorder — valor pedido y a la espera de stock (precio unitario × unidades en backorder)',
      tipAdj: 'Ajustado — Ventas año + Backorder (ventas potenciales si se despacha el backorder)',
      tipSP: 'MP {prev} — mismo período del año anterior: 1 de enero a la fecha de corte',
      tipDSP: 'Δ MP — diferencia frente al mismo período de {prev}',
      tipPSP: '% MP — variación porcentual frente al mismo período de {prev}',
      tipFY: 'AC {prev} — año completo {prev}: los 12 meses del año anterior',
      tipDFY: 'Δ AC — diferencia frente al año completo {prev}',
      tipPFY: '% AC — variación porcentual frente al año completo {prev}',
      showN: 'Mostrar', showAll: 'Todos', noItems: 'Sin productos para el período seleccionado.',

      foot: 'Panel ejecutivo · Fuente: <b>{source}</b> · datos al {date}. La cartera de pedidos se descompone en <b>pedido = por despachar (comprometido) + en backorder + ya despachado</b>. Todos los valores se calculan a partir del archivo cargado.'
    },
    en: {
      title: 'MRO International — Executive Summary',
      subtitle: 'International sales and order backlog · net invoiced amount',
      asOfBig: 'As of {date}',
      asOfMut: 'Year to date: {start} – {date}',
      reupload: 'Change file',
      share: 'Share', shareBuilding: 'Building…', shareCopied: '✓ Link copied', shareErr: 'Build failed',
      loadingShared: 'Loading shared data…',
      errShared: 'The shared link is invalid or corrupted.',

      upTitle: 'Load data file',
      upSub: 'Drag an <b>.xlsx</b> file here, or click to choose one.',
      upBrowse: 'Choose file',
      upFmt: 'The file must contain the sheets <code>Sales</code>, <code>Target</code>, <code>PendingFullFill</code> and <code>IF</code>, using the same column layout as the source report. All processing happens in your browser — nothing is uploaded.',
      errNoSheet: 'Required sheet not found: {sheet}.',
      errParse: 'Could not read the file. Make sure it is a valid .xlsx in the expected format.',
      errNoSales: 'The Sales sheet has no valid sales rows.',

      s1h: 'Sales — this year vs. last year',
      s1st: 'Net sales year to date, compared with the same period in {prev}',
      s1curLabel: 'Net sales {year} · year to date',
      s1curNote: '{start} – {date}, {year} · <b>{full}</b>',
      s1prevLabel: 'Same period {prev}',
      s1prevNote: '{start} – {date}, {prev}',
      s1custLabel: 'Active customers',
      s1custNote: 'across <b>{n}</b> countries · YTD',
      pillVs: '{arrow} {sign}{pct}  ({dsign}{delta} vs. {prev})',
      legMonthly: 'Monthly net sales · {month} excluded (partial month)',
      legCumNote: 'Cumulative net sales in the year · are we ahead or behind?',
      legCumCur: '{year} cumulative',
      legCumPrev: '{prev} cumulative',

      s2h: 'Markets & products',
      s2st: 'Where sales come from · {year} year to date',
      s2ctryLabel: 'Sales by country',
      s2ctryNote: 'Top international markets · ≈ {pct} of sales',
      s2segLabel: 'Product mix',
      s2segNote: 'Sales by line · {top} ≈ {pct} of total',

      s3h: 'Open sales orders',
      s3st: 'Orders booked and not yet fully shipped',
      s3valLabel: 'Open order value',
      s3valNote: '<b>{n}</b> open orders · <b>{m}</b> customers',
      s3compLabel: 'Order backlog composition',
      s3ready: 'To ship (committed)',
      s3bo: 'On backorder',
      s3units: 'Units ordered',

      s4h: 'Backorders',
      s4st: 'Ordered but awaiting stock',
      s4valLabel: 'Backorder value · awaiting stock',
      s4units: 'Units',
      s4sos: 'Orders affected',

      s5h: 'Awaiting fulfillment',
      s5st: 'Committed and allocated — awaiting Item Fulfillment (IF)',
      s5valLabel: 'Ready to ship · committed value',
      s5unitsLabel: 'Units allocated',
      s5shareLabel: '% of open value',

      byWh: 'By warehouse',
      byCust: 'Top by customer',

      s6h: 'Product performance',
      s6st: 'How products are selling · ranked by YTD sales — toggle backorder to see the impact',
      perfView: 'View', perfActual: 'Actual sales', perfBO: 'With backorder',
      tblSeg: 'Segment comparison',
      tblProd: 'Product comparison',
      abbrPerf: '<b>SP</b> = same period {prev} · <b>FY</b> = full year {prev} · <b>YTD sales</b> = year to date {year}',
      abbrPerfBO: '<b>Adjusted</b> = YTD sales + Backorder (unit price × units on backorder) · re-ranked by Adjusted',
      colProduct: 'Product', colSegment: 'Segment', colYTD: 'YTD sales', colUnits: 'Units',
      colSP: 'SP {prev}', colDSP: 'Δ SP', colPSP: '% SP', colFY: 'FY {prev}', colDFY: 'Δ FY', colPFY: '% FY',
      colBO: 'Backorder', colAdj: 'Adjusted',
      tipRank: 'Rank position',
      tipProduct: 'Product code (SKU) and name',
      tipSegment: 'Product line or segment',
      tipYTD: 'YTD sales — net sales year-to-date, from January 1 to the as-of date',
      tipUnits: 'Units sold during the period',
      tipBO: 'Backorder — value ordered and awaiting stock (unit price × units on backorder)',
      tipAdj: 'Adjusted — YTD sales + Backorder (potential sales if the backorder ships)',
      tipSP: 'SP {prev} — same period last year: January 1 to the as-of date',
      tipDSP: 'Δ SP — difference vs. the same period in {prev}',
      tipPSP: '% SP — percentage change vs. the same period in {prev}',
      tipFY: 'FY {prev} — full year {prev}: all 12 months of the prior year',
      tipDFY: 'Δ FY — difference vs. full year {prev}',
      tipPFY: '% FY — percentage change vs. full year {prev}',
      showN: 'Show', showAll: 'All', noItems: 'No products for the selected period.',

      foot: 'Executive panel · Source: <b>{source}</b> · data as of {date}. The order backlog breaks down as <b>ordered = to ship (committed) + on backorder + already shipped</b>. All values are computed from the uploaded file.'
    }
  };

  // Month names for chart axes and copy.
  const MONTHS = {
    es: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
    en: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
  };
  const MONTHS_LONG = {
    es: ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
    en: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
  };

  // ---- Label dictionaries: raw data value -> {es, en}. Fallback = raw value. ----
  const SEGMENTS = {
    '4T Engine Oils': { es: 'Aceites 4T', en: '4T Engine Oils' },
    '2T Engine Oils': { es: 'Aceites 2T', en: '2T Engine Oils' },
    'Aerosols': { es: 'Aerosoles', en: 'Aerosols' },
    'Suspension Oils': { es: 'Aceites suspensión', en: 'Suspension Oils' },
    'Maintenance': { es: 'Mantenimiento', en: 'Maintenance' },
    'Trans & Gear Oils': { es: 'Trans. y engranajes', en: 'Trans & Gear Oils' },
    'Merchandise': { es: 'Merchandising', en: 'Merchandise' },
    'Oil Filters': { es: 'Filtros de aceite', en: 'Oil Filters' },
    'Air Filters': { es: 'Filtros de aire', en: 'Air Filters' },
    'Grease': { es: 'Grasas', en: 'Grease' },
    'Coolant': { es: 'Refrigerantes', en: 'Coolant' },
    'Additives': { es: 'Aditivos', en: 'Additives' },
    'Cleaners': { es: 'Limpiadores', en: 'Cleaners' }
  };
  const WAREHOUSES = {
    'EU Warehouse': { es: 'Bodega EU', en: 'EU Warehouse' }
    // MXWHS, VTNWH-MX, ProWHS, etc. are codes — shown as-is.
  };
  const STATUS = {
    'Pending Fulfillment': { es: 'Pendiente de despacho', en: 'Pending fulfillment' },
    'Partially Fulfilled': { es: 'Despacho parcial', en: 'Partially fulfilled' },
    'Pending Billing': { es: 'Pendiente de facturación', en: 'Pending billing' },
    'Pending Billing/Partially Fulfilled': { es: 'Fact. pend. / despacho parcial', en: 'Pending billing / partially fulfilled' }
  };
  const COUNTRIES = {
    'Mexico': { es: 'México', en: 'Mexico' },
    'Greece': { es: 'Grecia', en: 'Greece' },
    'Taiwan (Province of China)': { es: 'Taiwán', en: 'Taiwan' },
    'China': { es: 'China', en: 'China' },
    'United States': { es: 'Estados Unidos', en: 'United States' },
    'United States of America': { es: 'Estados Unidos', en: 'United States' },
    'Spain': { es: 'España', en: 'Spain' },
    'Italy': { es: 'Italia', en: 'Italy' },
    'France': { es: 'Francia', en: 'France' },
    'Germany': { es: 'Alemania', en: 'Germany' },
    'Canada': { es: 'Canadá', en: 'Canada' },
    'Brazil': { es: 'Brasil', en: 'Brazil' },
    'Viet Nam': { es: 'Vietnam', en: 'Vietnam' },
    'Thailand': { es: 'Tailandia', en: 'Thailand' },
    'Japan': { es: 'Japón', en: 'Japan' },
    'Turkey': { es: 'Turquía', en: 'Turkey' },
    'Netherlands': { es: 'Países Bajos', en: 'Netherlands' },
    'Sweden': { es: 'Suecia', en: 'Sweden' },
    'New Zealand': { es: 'Nueva Zelanda', en: 'New Zealand' },
    'South Africa': { es: 'Sudáfrica', en: 'South Africa' },
    'Poland': { es: 'Polonia', en: 'Poland' },
    'Peru': { es: 'Perú', en: 'Peru' },
    'Korea, Republic of': { es: 'Corea del Sur', en: 'South Korea' },
    'Russian Federation': { es: 'Rusia', en: 'Russia' },
    'United Kingdom': { es: 'Reino Unido', en: 'United Kingdom' },
    'United Kingdom of Great Britain and Northern Ireland': { es: 'Reino Unido', en: 'United Kingdom' }
  };

  function look(dict, value, lang) {
    const v = String(value == null ? '' : value).trim();
    const hit = dict[v];
    return hit ? (hit[lang] || hit.en || v) : v;
  }

  root.I18N = {
    STR, MONTHS, MONTHS_LONG,
    seg: (v, lang) => look(SEGMENTS, v, lang),
    wh: (v, lang) => look(WAREHOUSES, v, lang),
    status: (v, lang) => look(STATUS, v, lang),
    country: (v, lang) => look(COUNTRIES, v, lang)
  };
})(window);
