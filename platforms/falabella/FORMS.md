# Downloaded Forms — Catalogue (Falabella Chile)

Every form this platform needs from the Falabella Seller Center to build **last
month's** financial report: where it is, how to get it, and what is inside. Same
structure as `platforms/mercadolibre/FORMS.md`.

**This document grows.** Keep the "verified" date honest.

- Stores used: `DAB_CL-Fala (Coco)` (seller "Hogaria", id `SC2C911`) — full month
  2026-09, all 7 forms; `AIS_CL-Fala` (seller "Tecnohub", `SC7CB30`) — first
  discovery, a very small store (5 orders in September)
- Ziniao platform name: `falabella-智利`
- Source of the routes: the owner's intake workbook (7 forms with screenshots,
  taken from an English UI — **the stores run in Spanish; use the screenshots only
  as a map**), then probed live
- Last verified: **2026-10-16**, by hand-driven probes. No download code exists yet.
- Host: `https://sellercenter.falabella.com` (one host for CL/PE/CO; the payments
  page has a CHILE / PERU / COLOMBIA switch, CHILE selected)

---

## Index — the owner's 7 forms

| # | Owner's name | Falabella name | Page | Period | Ready after | File |
|---|---|---|---|---|---|---|
| 1 | order | Exportar órdenes (Filtrado) | Órdenes → Lista de Órdenes | order date, **max 14 days** per export → 3 exports per month | 15–50 s, direct download | `export.xlsx` (rename it) |
| 2 | billing | Transacciones | Pagos → Órdenes y Transacciones | transaction date, the calendar month | ~20 s, async list | `NewReportTransaction_<ts>.xlsx` |
| 3 | stock | Todos los productos | Fulfillment → Productos | none — snapshot | ~3 s, async list | `Product_details_<id>.xlsx` |
| 4 | billing1 | Factura falabella.com report | Facturación → Factura falabella.com | billing period **21st → 20th** | instant, async list | `InvoiceReport_FACL-<rut>-<from>-<to>_<ts>.xlsx` |
| 5 | billing2 | Liquidación factura report | Facturación → Liquidación factura | calendar month | instant, async list | `<seller>_FSC_LIQ_BOLETA_NCR_REPORT_<from>_<to>.xlsx` |
| 6 | invoice1 | Documento tributario (PDF) | detail of a billing1 period, tabs Factura / Notas de crédito | billing period | immediate download | `<folio>_FACTURA.pdf`, `<folio>_NOTA_DE_CREDITO.pdf` |
| 7 | invoice2 | Liquidación PDF | detail of a billing2 month | calendar month | immediate download | `<uuid>.pdf` |

DAB, September 2026: orders 798 + 711 + 113 lines, transactions 10 090 lines,
stock 257 products, billing1 (21 Aug–20 Sep) 5 012 lines, billing2 200 lines,
invoice1 one factura + one nota de crédito, invoice2 one liquidation PDF.

**Timing for last month:** forms 1, 2, 3, 5, 7 are complete from the 1st of the
next month. The billing period 21 Sep–20 Oct (form 4/6) only appears after it
closes on the 20th, so September's billing1/invoice1 are complete only from ~21
October. The report decides what to do with that (see README, accounting rules);
the download records which periods were listed.

Also available, not one of the 7: Pagos → "Órdenes" (`ReportOrderSummary_<ts>.xlsx`,
English headers, only orders with money movement) — not needed now.

## How the pages are built — read before writing selectors

- **The new pages run inside shadow roots.** Orders, payments and invoicing are
  micro-frontends mounted under `div#orders`,
  `div#accountstatements-app-container-wrapper` and
  `div#settlement-liquidation-invoice-container`. `document.querySelector` and
  `document.body.innerText` see none of their content. Every lookup must walk
  open shadow roots; Selenium can click and type into elements returned from such
  a walk.
- **Stable hooks exist; prefer them to text:** `data-testid` values such as
  `export-btn-id`, `txn-search-btn`, `search-btn`, `fileType-export-btn-id`,
  `select-options-id`, `range-picker-id`, `settlements-download`,
  `liquidation-invoice-download-report`, `liquidation-invoice-check-details-btn`,
  `test_dropdown-header`, `test_list`; menu items carry `data-menu-id`
  (`…-ordersummary`, `…-alltransactionsV2`); tabs carry ids
  (`tab-5` = Todas on the orders page, `omnibar-search-tabs-tab-1` = Fecha de orden,
  `rc-tabs-0-tab-liquidation-invoice`); calendar cells carry `title="YYYY-MM-DD"`.
- **Language gate:** the payments/invoicing root
  `[data-testid=main-container-wrapper]` has `data-localefull="es-CL"` (and
  `data-seller`, `data-sellername`). Check it before downloading — column names of
  forms 1–5 are Spanish and follow the UI language. The footer has an `Idioma`
  switch (`Español`).
- **Pop-ups that block clicks:**
  - a guided tour on the orders page ("Nuevo · 1 de 5 · Tus envíos de hoy",
    button `Omitir`) — first visit per store;
  - a feedback modal on the payments page ("Qué te pareció la nueva pestaña de
    órdenes y transacciones?", button `Cancelar`) — it appears **with a delay**, so
    check for it before every click, not only after loading;
  - the dashboard survey (`#dismiss`);
  - success notifications on the invoicing page covering the row buttons.
- **Calendar days need real mouse clicks.** A JavaScript `element.click()` on an
  Ant Design calendar cell does nothing; use a real click (ActionChains). Buttons
  and menu items accept JavaScript clicks (useful when a notification covers them).
- **The dashboard is the old layout.** `/fulfillment/product` typed directly shows
  "Lo sentimos, tuvimos un problema técnico"; open `/index/dashboard` first, then
  `/fulfillment`.
- **Async forms share one "downloads" table pattern:** the request adds a row
  (`Procesando` → `Listo`; `Finalizado` on Fulfillment); `Actualizar` refreshes it;
  the row's `[data-testid=settlements-download]` button (a plain link on
  Fulfillment) downloads it. Rows stay listed for weeks, so an interrupted run can
  collect later by file name — the playbook's "owed" mechanism fits directly.
- Dates in the UI are Chile local time; file names mix UTC (`…T07:24:06.469Z`)
  and local (`…T04:24:14…`).

---

## 1. order — Lista de Órdenes export

Owner's note: download on the 2nd.

Route (verified on DAB 2026-10-16):
1. `/order`; wait ~14 s; close the guided tour (`Omitir`) if shown.
2. Tab `Todas` (`#tab-5` in the shadow root).
3. **Click inside the search box** "Buscar por número(s) de orden": a panel opens
   with filter tabs (Orden Nº, Fecha de orden, Fecha límite, Fulfillment, Bodega,
   Op. logístico, Documento, Multiparte). This is where the date filter lives; the
   page has no other date control.
4. Tab `Fecha de orden` (`#omnibar-search-tabs-tab-1`) → range picker with inputs
   `Fecha inicial` / `Fecha final` (`input[date-range=start|end]`); presets Hoy,
   Últimos 3/7/14 días. Click the start day, then the end day (real clicks; navigate
   with `.ant-picker-header-prev-btn` / `-next-btn`). **After the start day is
   chosen, only the next 13 days are selectable** — a month needs three exports:
   1–14, 15–28, 29–end.
5. The chip "sep 1, 2026 - sep 14, 2026" appears in the search bar and the list
   loads; the end day **is included** (orders up to 14 Sep 23:39 shown).
6. `Exportar` (enabled now) → `Filtrado` → modal "Exportaciones" → `Descargar` →
   `XLSX`. The browser downloads `export.xlsx` directly after 15–50 s (no list);
   rename it at once, three exports would overwrite each other.

On a store with no orders in the range `Exportar` stays disabled (seen on
AIS_CL-Fala) — that is "no data", not a failure.

File: sheet `Sheet 1`, header row 1, 69 columns, one row per order line (`Item
id`): `SKU del vendedor, SKU de Falabella, Fecha de creacion ("Sep 14, 2026
23:39"), N° orden, Orden id, Origen del pedido, Moneda, Documento requerido,
…customer and address columns…, Metodo de pago, Precio pagado, Precio unitario,
Costo de envio, Cupon/Descuentos, Impuesto ("13,990.00" text), Producto,
Variacion, Operador Logistico, Tipo de envio, …tracking…, Fecha limite de
despacho, Estado (Entregadas / Canceladas / Devoluciones / Enviados / Listo para
despachar), Tipo de Fulfillment, Motivo de cancelacion, …billing-document
columns`. One column is literally named `[object Object]` (a Falabella bug).

**Contains customer personal data** (name, RUT, address, phone, email). It lives
only under `downloads/`; never in git, never in test fixtures.

## 2. billing — Transacciones (the money file)

Owner's note: download on the 5th. The owner's earlier rule "5 extra days each
side" is **not needed**: on DAB, 1–30 Sep returned exactly 1–30 Sep (302 lines on
the 30th, nothing outside). Use the calendar month.

Route (verified on DAB and AIS 2026-10-16):
1. `/report/transaction/statement#/ordersandtransactionsoverview`; wait ~25 s;
   close the feedback modal whenever it appears.
2. Filter select (`.ant-select-selector`, first on the page) → `Fecha transacción`.
3. Click the read-only input "Seleccione el período de la transacción…" → custom
   two-month calendar `[data-testid=range-picker-id]` (presets Hoy, Últimos 7/15/30
   días, Personalizado; arrows `.anticon-left` / `.anticon-right`). The 1st of the
   current month is pre-marked: **clicking an already-marked day does not start a
   range** — if `Aplicar` stays disabled after two clicks, click the start day
   again. `Aplicar`, then read the input back:
   `"1  septiembre   2026 - 30  septiembre   2026"`.
4. **Press the search button** (`button.search-btn`, testid `txn-search-btn`);
   until then the "Transacciones" menu item is disabled.
5. `[data-testid=export-btn-id]` "Descargar reporte" → `Transacciones`
   (`data-menu-id` ends in `alltransactionsV2`).
6. Modal "Descargar reporte de transacciones": format select
   `[data-testid=select-options-id]` → `XLSX`; `[data-testid=fileType-export-btn-id]`.
   Notice: "La creación del reporte puede tomar hasta 3 minutos."
7. "Documents descargados": new row `Procesando` → `Listo` after ~20 s (DAB, 10 090
   lines); download with the row's `settlements-download` button.

File: sheet `Reporte Transacciones`, rows 1–4 banner ("Fecha de actualización",
"Este reporte podría tener cambios hasta el cierre del período de facturación"),
row 5 group headers, **row 6 headers**, data from row 7. 31 columns:

`Fecha creación de la orden, Nombre del producto, Nº de orden, SKU vendedor, SKU
Falabella, Id Artículo, Falabella-Id, Estado de la orden, Fecha de Entrega, Tipo
de transacción, Fecha de transacción, Divisa, Monto (Sin IVA), IVA, Monto con IVA,
Tipo de comisión, % comisión, Talla, Opscore, Estado de pago, Nº estado de cuenta,
Referencia de pago Fpay, Comentarios, Tipo de envío, Proveedor logístico,
Modalidad logística, Seller ID, Tipo de Seller, Tipo de Documento Tributario,
N° Documento Tributario, Categoría de transacciones`

Second sheet `Glosario de Terminos Contables`: 147 transaction types with a
description — the starting point of the fee catalogue (some descriptions have
broken accents, e.g. `env√≠o`).

One line per money movement per order line: the sale is `Pago por precio del
producto` (category `Pesos`), fees are negative (`Cobro por comisión por venta`,
`Cobro por cofinanciamiento logístico`, …), buyer-paid shipping comes as a pair
`Pago de envío comprador` / `Reversa de pago de envío comprador`. Store-level
charges without an order (sponsored products, FBF storage) also appear.
`Nº estado de cuenta` is the payout statement (`FACL-<seller>-20260916-CLP`);
`Estado de pago` is `Pagado` / `No Pagado`.

## 3. stock — Fulfillment products

Owner's note: two moments — local midnight on the 1st, and Mondays 15:00.

Route: `/index/dashboard` (close `#dismiss`), then `/fulfillment` (lands on
`/fulfillment/product`). Dropdown `[data-testid=test_dropdown-header]` with text
`Descargar` → `[data-testid=test_list] a` → `Todos los productos`. A row appears
at the bottom in "Exportaciones" ("Products Export", `Finalizado` within seconds,
newest first); its link points to
`storage.googleapis.com/fullfillment-by-storage-production/fby-exports/Product_details_<id>.xlsx`.
Scroll the link into view, then a real click downloads it. Other tabs (Solicitudes,
Retiros, Movimientos de Inventario, Colecciones, Desempeño) are not needed yet.

File: sheet `Product details`, header row 1: `Marca, Categoria (JSON), Codigo
Universal, Nombre, SKU Vendedor, SKU GSC, Barcode Etiqueta de Producto, Modelo,
Color, Tamano, Precio, Ventas ultimas 4 semanas, Stock Disponible`. A snapshot: it
cannot be produced for a past date, hence the owner's fixed times.

## 4. billing1 — Factura falabella.com report

Owner's note: download on the 10th.

Route: `/invoicing#/settlement-invoice?period=2026` (year select `Fecha`; in
January the previous year must be chosen). Table
`[data-testid=settlement-invoice-table]`, one row per billing period **21st →
20th** ("21 agosto 2026 - 20 septiembre 2026"), columns Facturas, Notas de crédito,
Notas de crédito/débito promociones. In the row: click the `Descargar reporte`
dropdown trigger (`[class*=dropdown-trigger]`, a click, not hover) → `CSV` /
`XLSX`. The file is listed in "Descargas" at once (generated when the period
closed, only re-listed on request); download with `settlements-download`.

A calendar month touches two billing periods (September: 21 Aug–20 Sep and 21
Sep–20 Oct). A period is listed only after it closes.

File: sheet `settlement invoice`, header row 1, 22 columns: `Numero de
documento, Tipo de documento (Factura / Nota de Credito), Fecha de Transaccion,
Descripcion Factura, Tipo de Transaccion, Nombre del producto, SKU vendedor, SKU
Falabella, Monto (Sin IVA), IVA, Monto con IVA, Divisa, N estado de cuenta, N de
orden, Referencia de pago Fpay, Falabella-Id, Estado de la orden, Modalidad
logistica, Tipo de envio, Detalles de transaccion, Seller ID, Tipo de Seller`.
Lines match the transaction types of form 2 and carry the same statement and
order numbers — the bridge between invoiced fees and paid-out fees.

## 5. billing2 — Liquidación factura report

Owner's note: download on the 10th, "last period".

Route: same page, tab `Liquidación factura`. One row per **calendar month**
("01 septiembre 2026 - 30 septiembre 2026"), columns Boletas, Notas de crédito,
Monto total liquidado (DAB September: 28.813.590 − 1.060.227 = 27.753.363 CLP).
Row button `[data-testid=liquidation-invoice-download-report]` → listed at once in
"Descargas" as `FSC_LIQ_<from>_<to>.xlsx`; the downloaded file is named
`<seller>_FSC_LIQ_BOLETA_NCR_REPORT_<from>_<to>.xlsx`.

File: sheet `Liquidaciones` (seller, period, folio, referenced doc type, SKU,
article, quantity, line total, liquidation total) and sheet `Docs. liquidados`
(issue time, document type `BOLETA` / `FACTURA`, folio, order number,
Falabella-Id, SKU, quantity, line amount, document total, referenced doc). These
are the buyer documents Falabella issued in the seller's name.

## 6. invoice1 — fee invoice PDFs

Detail page of a billing1 period (`Revisar detalle` in the row →
`#/details/settlement/<uuid>/invoice`): totals by group (DAB 21 Aug–20 Sep and AIS
the same period checked) and tabs Factura / Notas de crédito / Documentos de
promociones. On each tab, `Descargar documento tributario (PDF)` opens a menu
with one item per folio (`520878`, `505871`); each item's `data-menu-id` holds the
file URL (`storage.googleapis.com/seller-invoices-production/settlement-invoices/facl/<folio>_FACTURA.pdf`).
A real click on the item downloads it (a `fetch()` of the URL is blocked by CORS).
"Documentos de promociones" had no PDF on either store.

## 7. invoice2 — liquidation PDFs

Detail page of a billing2 month (`[data-testid=liquidation-invoice-check-details-btn]`
→ `#/details/liquidation/<id>`): one row per liquidation document (DAB September:
`41028 Global 27.753.363 CLP Emitida`) → `Descargar` → `PDF`; the file is named
by a UUID — rename it to the folio.

The PDFs are for the accountant; the report reads the XLSX forms.

---

## Numbers measured (playbook §8)

| What | DAB_CL-Fala (Coco), September 2026 |
|---|---|
| order, per 14-day export | 50 s (798 lines), 39 s (711), 13 s (113) |
| billing (Transacciones) request → ready | ~22 s (10 090 lines) |
| stock export | seconds |
| billing1 / billing2 | listed instantly |
| PDFs | immediate |
| store open (first time on the VM) | 2 retries of Ziniao −10000 (kernel download), then fine |

Still to measure: one unattended run end to end, range limit of the payments
calendar, how long the downloads tables keep old files.
