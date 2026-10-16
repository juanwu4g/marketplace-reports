# Downloaded Forms — Catalogue (Falabella Chile)

Every form this platform needs from the Falabella Seller Center: where it is,
how to get it, and what is inside. Same structure as
`platforms/mercadolibre/FORMS.md`.

**This document grows.** Keep the "verified" date honest.

- Store used for discovery: `AIS_CL-Fala` (Ziniao platform `falabella-智利`,
  seller name "Tecnohub", seller id `SC7CB30`)
- Source of the routes: the owner's intake workbook (6 forms with screenshots),
  then probed live on 2026-10-16
- Last verified: **2026-10-16**, each form fetched once by hand-driven probes.
  No download code exists yet.
- Host: `https://sellercenter.falabella.com` (one host for CL/PE/CO; the country
  is a switch on the payments page, `CHILE` is selected for this seller)

---

## Index

| # | Form | Page | Period control | Ready after | File |
|---|---|---|---|---|---|
| 1 | Transacciones (orders and transactions) | Pagos → Órdenes y Transacciones | filter "Fecha transacción" + date range | ~20 s, async | `NewReportTransaction_<ts>.xlsx` |
| 2 | Órdenes (order summary) | same page | filter "Fecha de la orden" + date range | ~10 s, async | `ReportOrderSummary_<ts>.xlsx` |
| 3 | Stock (Fulfillment products) | Fulfillment → Productos | none, snapshot | ~3 s, async | `Product_details_<id>.xlsx` |
| 4 | Factura falabella.com report | Facturación → Factura falabella.com | billing period row (21st → 20th) | instant, listed | `InvoiceReport_FACL-<rut>-<from>-<to>_<ts>.xlsx` |
| 5 | Liquidación factura report | Facturación → Liquidación factura | billing period row (calendar month) | instant, listed | `SC7CB30_FSC_LIQ_BOLETA_NCR_REPORT_<from>_<to>.xlsx` |
| 6 | Tax documents (PDF) | detail page of 4 and of 5 | the same period rows | immediate download | `<folio>_FACTURA.pdf`, `<folio>_NOTA_DE_CREDITO.pdf`, `<uuid>.pdf` |

The intake workbook lists the orders page (`/order` → Exportar) for the sales
form. On this store that page has **no date control** (see form 2), so the order
summary from the payments page is used instead until the owner decides.

## How the pages are built — read before writing selectors

- **The new pages run inside shadow roots.** The orders, payments and invoicing
  pages are micro-frontends mounted under `div#orders`,
  `div#accountstatements-app-container-wrapper` and
  `div#settlement-liquidation-invoice-container`. `document.querySelector` and
  `document.body.innerText` see none of their content (a first probe saw only the
  menu). Every lookup must walk open shadow roots; Selenium can click and type into
  elements returned from such a walk.
- **Stable hooks exist; prefer them to text:** `data-testid` values such as
  `export-btn-id`, `txn-search-btn`, `search-btn`, `fileType-export-btn-id`,
  `select-options-id`, `range-picker-id`, `settlements-download`,
  `liquidation-invoice-download-report`, `liquidation-invoice-check-details-btn`;
  menu items carry `data-menu-id` (`…-ordersummary`, `…-alltransactionsV2`); tabs
  carry ids (`rc-tabs-0-tab-ordersandtransactionsoverview`,
  `rc-tabs-0-tab-liquidation-invoice`); calendar cells carry `title="YYYY-MM-DD"`.
- **Language gate:** the micro-frontend root
  `[data-testid=main-container-wrapper]` has `data-localefull="es-CL"` (also
  `data-seller="SC7CB30"`, `data-sellername`). Check it before downloading: column
  names of forms 1, 4, 5, 3 are Spanish and follow the UI language (form 2's are
  English regardless). The owner's screenshots are from an English UI, so some
  operators do switch language.
- **Pop-ups that block clicks:** a feedback modal on the payments page
  ("Qué te pareció la nueva pestaña…", close with its `Cancelar`), the dashboard
  survey (`#dismiss`), and success notifications on the invoicing page that cover
  the row buttons. Clicking through JavaScript (`element.click()`) avoids the
  last; the first must be closed.
- **The dashboard is the old layout** (top menu, jQuery). `/fulfillment/product`
  typed directly shows "Lo sentimos, tuvimos un problema técnico"; open
  `/index/dashboard` first, then `/fulfillment` (the owner's notes say the same).
- **Every async form shares one "downloads" table pattern:** the request adds a
  row (`Procesando` → `Listo`, or `Finalizado` on Fulfillment); `Actualizar`
  refreshes it; the row's `[data-testid=settlements-download]` button (a plain
  link on Fulfillment) downloads it. Requests stay listed for weeks (13 pages of
  history on the payments page), so a run that was interrupted can collect later
  by file name — the playbook's "owed" mechanism fits directly.
- Dates in the UI are Chile local time; file names mix UTC (`…T07:24:06.469Z`)
  and local (`…T04:24:14…`).

---

## 1. Transacciones — the money file

Owner's note: download on the 5th of the next month; **select 5 extra days on
each side of the month, delete them after downloading.**

Route (verified 2026-10-16):
1. `/report/transaction/statement#/ordersandtransactionsoverview`; wait ~25 s
   (skeleton first); close the feedback modal if present.
2. Filter select (`.ant-select-selector`, first on the page) → option
   "Fecha transacción".
3. Click the read-only input "Seleccione el período de la transacción…" → custom
   two-month calendar `[data-testid=range-picker-id]`: navigate with `.anticon-left`
   / `.anticon-right`, click the start cell then the end cell
   (`td.ant-picker-cell-in-view[title=YYYY-MM-DD]`), `Aplicar`. Read the input back:
   `"27  agosto   2026 - 5  octubre   2026"`.
4. `[data-testid=txn-search-btn]` (optional; it shows the rows on screen).
5. `[data-testid=export-btn-id]` "Descargar reporte" → menu → "Transacciones"
   (`data-menu-id` ends in `alltransactionsV2`).
6. Modal "Descargar reporte de transacciones": format select
   `[data-testid=select-options-id]` → `XLSX`; `[data-testid=fileType-export-btn-id]`.
   Notice: "La creación del reporte puede tomar hasta 3 minutos."
7. Downloads table "Documents descargados": new row `Procesando`, `Listo` after
   ~21 s; download with the row's `settlements-download` button.

**Why the ±5 days:** the selected range is sent as local midnight to local
midnight — the disabled "Transacciones" item for the order-date filter showed
`Tue, 01 Sep 2026 04:00:00 GMT,Wed, 30 Sep 2026 03:00:00 GMT` for "1–30 Sep",
i.e. the end day is **not included**. Chile also changes UTC offset in early
September (04:00 → 03:00 GMT above). Pad the range, then keep rows by
`Fecha de transacción` inside the month in the report layer.

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

One line per money movement per order line: the sale itself is
`Pago por precio del producto` (category `Pesos`), fees are negative
(`Cobro por comisión por venta`, `Cobro por cofinanciamiento logístico`, …),
buyer-paid shipping comes as a pair `Pago de envío comprador` /
`Reversa de pago de envío comprador`. Store-level charges without an order
(sponsored products, FBF storage) also appear. `Nº estado de cuenta` is the
payout statement (`FACL-SC7CB30-20260916-CLP`); `Estado de pago` is
`Pagado` / `No Pagado`. AIS_CL-Fala, 27 Aug–5 Oct: 74 rows, 14 statements.

## 2. Órdenes — order summary

Same page and steps as form 1, but filter "Fecha de la orden" (input
"Seleccione el período de la orden…", search button `[data-testid=search-btn]`),
menu item "Órdenes" (`…-ordersummary`; it is disabled under the transaction-date
filter). Same format modal, ready in seconds, same downloads table.

File: sheet `ordersummary`, header row 1, **English headers**:
`Order No., Order Date, Transaction Date, Product, Falabella Id, Seller SKU, Price,
Commission, Fees, Payout Amount, Operational Status`. One row per order line.
AIS_CL-Fala September: 5 rows (`Delivered`, one `Returned` with price 0).

Open: it is not yet known whether orders with no money movement (cancelled
before payment, still in transit) appear here. The orders page would list them,
but on this store it cannot be exported (below).

### The orders page (`/order`) — not usable on this store

The owner's screenshot shows a date chip and a calendar icon next to the search
box. Live on 2026-10-16 the page shows "Debes seleccionar un filtro de búsqueda —
Ingresa un número de orden o selecciona un rango de fecha", but there is **no date
control** anywhere (all tabs checked, page zoomed out); the icon is now a cart and
does nothing; `Exportar` stays disabled. Either the page changed since the
screenshot or this seller has a different version. Re-check on another store.

## 3. Stock — Fulfillment products

Owner's note: two moments — local midnight on the 1st, and Mondays 15:00.

Route: `/index/dashboard` (close `#dismiss` survey), then `/fulfillment` (lands on
`/fulfillment/product`). Dropdown `[data-testid=test_dropdown-header]` with text
`Descargar` → list `[data-testid=test_list] a` → "Todos los productos". A row
appears in "Exportaciones" ("Products Export", `Finalizado` within ~3 s); its link
points to `storage.googleapis.com/fullfillment-by-storage-production/fby-exports/…`
and clicking it downloads the file. Other tabs: Solicitudes, Retiros, Movimientos
de Inventario (`/fulfillment/inventory-movement`), Colecciones, Desempeño — not
needed yet.

File: sheet `Product details`, header row 1: `Marca, Categoria (JSON), Codigo
Universal, Nombre, SKU Vendedor, SKU GSC, Barcode Etiqueta de Producto, Modelo,
Color, Tamano, Precio, Ventas ultimas 4 semanas, Stock Disponible`. A snapshot:
it cannot be produced for a past date, hence the owner's fixed times.

## 4. Factura falabella.com — fee invoice report

Owner's note: download on the 10th.

Route: `/invoicing#/settlement-invoice?period=2026` (year select `Fecha`). Table
`[data-testid=settlement-invoice-table]`, one row per billing period
**21st → 20th** ("21 agosto 2026 - 20 septiembre 2026"), columns Facturas, Notas
de crédito, Notas de crédito/débito promociones. In the row: hover-free click on
the `Descargar reporte` dropdown trigger → menu items `CSV` / `XLSX`. The file
appears at once in "Descargas" (it is generated when the period closes and only
re-listed on request); download with `settlements-download`.

A calendar month is covered by two billing periods (September = part of 21 Aug–20
Sep and part of 21 Sep–20 Oct); the second is listed only after it closes.

File: sheet `settlement invoice`, header row 1, 22 columns: `Numero de
documento, Tipo de documento (Factura / Nota de Credito), Fecha de Transaccion,
Descripcion Factura, Tipo de Transaccion, Nombre del producto, SKU vendedor, SKU
Falabella, Monto (Sin IVA), IVA, Monto con IVA, Divisa, N estado de cuenta, N de
orden, Referencia de pago Fpay, Falabella-Id, Estado de la orden, Modalidad
logistica, Tipo de envio, Detalles de transaccion, Seller ID, Tipo de Seller`.
Lines match the transaction types of form 1 and carry the same statement and
order numbers — the bridge between invoiced fees and paid-out fees.

The detail page (`Revisar detalle` → `#/details/settlement/<uuid>/invoice`) shows
the same totals by group (e.g. 21 Aug–20 Sep: total sin IVA 607.577, IVA 115.440,
con IVA 723.017 CLP) with tabs Factura / Notas de crédito / Documentos de
promociones.

## 5. Liquidación factura — sales documents issued for the seller

Owner's note: download on the 10th, "last period".

Route: same page, tab `Liquidación factura` (`#/liquidation-invoice`). One row
per **calendar month** ("01 septiembre 2026 - 30 septiembre 2026"), columns
Boletas, Notas de crédito, Monto total liquidado. Row button
`[data-testid=liquidation-invoice-download-report]` → file listed at once in
"Descargas" as `FSC_LIQ_<from>_<to>.xlsx`; the downloaded file is named
`SC7CB30_FSC_LIQ_BOLETA_NCR_REPORT_<from>_<to>.xlsx`.

File: sheet `Liquidaciones` (seller, period, folio, referenced doc type, SKU,
article, quantity, line total, liquidation total) and sheet `Docs. liquidados`
(issue time, document type `BOLETA` / `FACTURA`, folio, order number,
Falabella-Id, SKU, quantity, line amount, document total, referenced doc). These
are the buyer documents (boletas/facturas) Falabella issued in the seller's name;
September for AIS_CL-Fala: liquidation 41278, 99.440 CLP.

## 6. Tax documents (PDF)

- Invoice side: detail page of form 4 → each tab (Factura, Notas de crédito) has
  `Descargar documento tributario (PDF)`; its menu lists one item per document
  folio (`521806`, `465507`), each item's `data-menu-id` holds the file URL
  (`storage.googleapis.com/seller-invoices-production/settlement-invoices/facl/521806_FACTURA.pdf`).
  A real click on the item downloads it; a `fetch()` of the URL is blocked by CORS.
- Liquidation side: detail page of form 5 (`#/details/liquidation/<id>`) → one row
  per liquidation document → `Descargar` → `PDF`; the file is named by a UUID.

The PDFs are for the accountant; the report reads the XLSX forms.

---

## Numbers measured (playbook §8)

| What | AIS_CL-Fala, 2026-10-16 |
|---|---|
| Transacciones, request → ready | ~21 s for 27 Aug–5 Oct |
| Órdenes, request → ready | a few seconds |
| Stock export | ~3 s |
| Invoice / liquidation reports | listed instantly |
| Store open (first time on the VM) | 2 retries of Ziniao −10000 (kernel download), then fine |

Still to measure: one full run end to end, a large store, range limits of the
date picker, and how long the downloads tables keep old files.
