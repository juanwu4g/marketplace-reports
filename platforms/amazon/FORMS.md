# Downloaded Forms — Catalogue (Amazon Mexico)

Every form this platform needs from Amazon Seller Central (amazon.com.mx) to build
**last month's** financial report: where it is, how to get it, and what is inside.
Same structure as `platforms/falabella/FORMS.md`.

**This document grows.** Keep the "verified" date honest.

- Store used: `SM-Amazon` (display name "HOLAMX", legal name SHEN MING SA DE CV,
  RFC SMI120201D78) — full month 2026-09, all 5 forms. `PW-Amazon` exists in
  Ziniao but is idle: **never open it**.
- Ziniao platform name: `亚马逊-墨西哥`
- Source of the routes: the owner's intake workbook (5 forms with screenshots,
  taken from an English UI — the store runs in Spanish; use the screenshots only as
  a map), then probed live.
- Last verified: **2026-10-19**, by hand-driven probes in the test VM.
- Host: `https://sellercentral.amazon.com.mx`

---

## Index — the owner's 5 forms

| # | Owner's name | Amazon name | Page | Period | Ready after | File |
|---|---|---|---|---|---|---|
| 1 | order | Transacciones (Informe de transacciones) | Finanzas → Reportes de finanzas (`/payments/reports-repository`) | **calendar month** (option "Mensual"), Amazon's month in GMT-6 | ~45 s, async list | `<YYYY><Mon>MonthlyTransaction.csv` |
| 2 | summary | Resumen | same page, report type "Resumen" | calendar month | ~30–60 s, async list | `<YYYY><Mon>MonthlySummary.pdf` |
| 3 | invoice | Facturas de tarifas aplicables al vendedor | Tax Document Library (`/tax/seller-fee-invoices`) | settlement periods of ~14 days (e.g. Aug 31 → Sep 14) | already there; "Ver" opens a tab | CFDI `<uuid>.xml` + `<uuid>.pdf` |
| 4 | stock | Inventario de Logística de Amazon | FBA inventory (`/inventoryplanning/manageinventoryhealth`) → Reportes → Reporte de inventario → `/reportcentral/MANAGE_INVENTORY_HEALTH/1` | none — snapshot of today | seconds, async list | `<number>.csv` |
| 5 | taxcert | Certificado de retención | Facturas fiscales mexicanas (`/sellertaxinvoice/displaySellerTaxInvoices`) | calendar month | already there | `<uuid>_<n>.xml`, `<uuid>_<n>.pdf` |

SM-Amazon, September 2026: transactions 1 084 lines (Pedido 1 020, Reembolso 40,
Ajuste 10, fee lines 11, Trasferir 3), summary 1 page, fee CFDIs for 3 settlement
periods (2 invoices + 1 payment complement per period, each as PDF and XML),
stock 332 SKUs, withholding certificate with 1 024 service lines.

**Timing for last month:** forms 1, 2 and 5 are complete from the first days of
the next month (the September certificate was stamped 2026-10-06 — run on or after
the 7th). The fee invoice of the period that **straddles** month end (Sep 28 → Oct
12) appears only after that period closes; a run before then records it as
pending (same idea as Falabella's billing1).

## How the pages are built — read before writing selectors

- **Katal web components** (`kat-dropdown`, `kat-radiobutton`, `kat-button`,
  `kat-table-row`, `kat-statusindicator` …) with open shadow roots. Walk shadow
  roots for every lookup (`core/browser_kit`).
- **Setting a kat-dropdown:** set its `value` and dispatch a composed `change`:
  `e.value=v; e.dispatchEvent(new CustomEvent('change',{bubbles:true,composed:true,detail:{value:v}}))`.
  Clicking options is unreliable (options render lazily in the shadow root). The
  available values are in the property `element.options` (`[{name,value}]`).
- **Language gate:** the header `#ngstrim-header` carries `data-locale`
  (`es_MX` when Spanish; `<html lang>` is `es-MX`). Column names and the summary
  PDF follow the UI language. **Switching back** is a page load of the same URL
  with `mons_sel_locale=es_MX` (the header's language links are
  `?ref_=xx_swlang_head_xx&mons_sel_locale=<xx_>`; Español is not listed while
  it is the current language).
- **Flaky services:** the fee invoice page sometimes renders "No invoices
  available" / "Unable to connect to service" although invoices exist (seen
  2026-10-19, 250 rows an hour earlier). That message is **never** NO_DATA — retry
  with waits, then PAGE_BLOCKED.
- **Other users:** the store is used by people (inventory reports requested on
  18–19 Oct by someone else). Only request reports; never change settings.
- **Captchas / login:** if a sign-in or captcha page appears, stop the store
  and hand it to a person (需要人工处理). Never solve or bypass it.

## 1. order — Transacciones (monthly)

Page `/payments/reports-repository`. Controls:

- report type `kat-dropdown` values: `SELLER_SUMMARY_DATE_RANGE` (Resumen),
  `SELLER_TRANSACTION_DATE_RANGE` (Transacciones), `SELLER_HOLD_DATE_RANGE`
  (Transacción aplazada — not needed: every line of the transaction report carries
  "Estado de la transacción", which shows deferral)
- period `kat-radiobutton` values `CUSTOM` / `MONTHLY` — click the inner label in
  its shadow root
- month `kat-dropdown`, **0-based** (`"8"` = Septiembre), year `kat-dropdown`
- request button `kat-button#filter-generate-button` (click the `button` in its
  shadow root)
- the list: `kat-table-row`s; status in `kat-statusindicator[label]` ("En curso",
  "Listo", "Vencido"); action `kat-button[label]` ("Actualizar" while running,
  "Descargar CSV" / "Descargar PDF" when ready, "Solicitar de nuevo" when expired).
  A row's text shows the requested range: `1 sept 2026, 0:00:00 GMT-6 30 sept 2026,
  23:59:59 GMT-6` — match the row by that range **and** the type, newest first.

The CSV: UTF-8 with BOM, **9 preamble lines** (title, account, range, notes),
then the header. Spanish headers:

`fecha/hora` ("31 ago 2026 11:28:19 p.m. GMT-7"), `Id. de liquidación`, `tipo`,
`Id. del pedido`, `sku`, `descripción`, `cantidad`, `marketplace`,
`cumplimiento`, `ciudad del pedido`, `estado del pedido`,
`código postal del pedido`, `modelo de recaudación de impuestos`,
`ventas de productos`, `impuesto de ventas de productos`, `créditos de envío`,
`impuesto de abono de envío`, `créditos por envoltorio de regalo`,
`impuesto de créditos de envoltura`, `Tarifa reglamentaria`,
`Impuesto sobre tarifa reglamentaria`, `descuentos promocionales`,
`impuesto de reembolsos promocionales`, `impuesto de retenciones en la plataforma`,
`tarifas de venta`, `tarifas fba`, `tarifas de otra transacción`, `otro`, `total`,
`Estado de la transacción`, `Fecha de liberación de la transacción`.

Numbers use `,` as thousands separator. Rows are printed in GMT-7 although the
month is Amazon's GMT-6 month: **use the report's month, never refilter by the
printed time**.

`tipo` values seen (Aug + Sep 2026): Pedido, Reembolso, Ajuste, Tarifa de
inventario FBA, Tarifa de servicio (Costo de la publicidad, Suscripción), Tarifas
de transacciones de Logística de Amazon (almacenamiento, almacenamiento
prolongado, retiro/puesta a disposición, envío con transportista asociado),
Trasferir (payout "Para la cuenta que termina en 460"). Facts that matter:

- `ventas de productos` is **net of IVA**; `impuesto de ventas de productos` is
  exactly 16 %.
- `impuesto de retenciones en la plataforma` is the **IVA withholding** (50 % of
  the IVA, i.e. 8 %) per order.
- **ISR withholding** is one `Ajuste` line without description around the 11th,
  for the **previous** month (Sep 11: −11 814.82 = August's ISR). A +2.69 line
  without description is "Ajuste del IVA del producto".
- Advertising: the `Tarifa de servicio` line "Costo de la publicidad" carries the
  net cost in `tarifas de otra transacción` and its IVA in `otro`
  (Sep: 17 178.35 + 2 748.54); "Suscripción" 600.00 in `otro`.
- Referral fees (`tarifas de venta`) are exact rates (10/12/14/15/18 %) of the
  IVA-inclusive price.
- September `total` = −144 640.43 = the summary's net.

## 2. summary — Resumen (monthly PDF)

Same page, type `SELLER_SUMMARY_DATE_RANGE`, period MONTHLY. One page;
`pdftotext -layout` reads it. Lines: Ingresos / Gastos / Transferencias /
Impuesto totals, then debit/credit detail (Ventas de producto FBA, reembolsos,
Créditos de envío, Descuentos promocionales, Tarifas de venta FBA, Tarifas de
transacción FBA, Tarifas de inventario y de servicios de Logística de Amazon,
Tarifas de servicio, Tarifas de administración de reembolso, Ajustes, Costo de la
publicidad …; Impuesto: retenidos, reembolsados, IVA retenido, ISR retenido,
Ajuste del IVA). Used as an **independent control total** for form 1.

## 3. invoice — Seller fee CFDIs

Page `/tax/seller-fee-invoices` (Tax Document Library, tab "Facturas de tarifas
aplicables al vendedor"). A plain `<table>`; each `TR` reads
`<Invoice Type> <PDF|XML> <Invoice Number (UUID)> <payer> <supplier> Amazon.com.mx
<Start Date> <End Date> Ver`, dates like `Mon Sep 14 19:07:36 UTC 2026`. Up to
1 000 most recent; "Load More" beyond.

Types: "Factura de la tarifa de vendedor para México" (two per period:
`TipoDeComprobante` **I** = fees, **E** = credit note "Ajuste a tarifas de
vendedores" related by `CfdiRelacionados`), "Complemento de pago de la tarifa de
vendedor para México" (start = end = payment day), occasionally "Spft Settlement
Note" / "Spft PaymentComplement Note".

**Download:** "Ver" opens a new tab at
`/documents/download/<uuid>/<invoice-number>.<xml|pdf>`; nothing lands in the
download folder. Read that URL from the tab, close the tab, then fetch it from
the main page with `fetch(url, {credentials:'include'})` and save the bytes
(verified: XML 200 `application/xml`, PDF 200 `application/pdf`).

Which periods belong to a month: every invoice whose period overlaps the month
(Sep 2026: Aug 31→Sep 14, Sep 14→Sep 28, Sep 28→Oct 12) plus the payment
complements dated inside the month.

## 4. stock — Inventario de Logística de Amazon

From `/inventoryplanning/manageinventoryhealth`: button "Reportes" → "Reporte de
inventario" opens a **new tab** `/reportcentral/MANAGE_INVENTORY_HEALTH/1`
(go there directly). Buttons "Solicitar descarga de .csv" / ".txt"; the list is
`kat-table-row`s with cells Tipo de reporte, Rango de fechas cubierto (request
time "19/10/26, 2:07 a.m."), Fecha solicitada, Formato de archivo, and a
"Descargar" button. Ready within seconds; the file lands normally.

CSV (UTF-8 BOM, English headers whatever the UI language): `snapshot-date`, `sku`,
`fnsku`, `asin`, `product-name`, `condition`, `available`, `inv-age-*` buckets,
`units-shipped-t7/t30/t60/t90`, `your-price`, `estimated-excess-quantity`,
`estimated-storage-cost-next-month`, `inbound-*`, `Reserved FC Processing`,
`Reserved Customer Order`, … SM-Amazon 2026-10-19: 332 SKUs, 2 583 available.

A snapshot of the run day, like Falabella's stock form.

## 5. taxcert — Certificado de retención (monthly)

Page `/sellertaxinvoice/displaySellerTaxInvoices` (Facturas fiscales mexicanas).
Rows `TR#invoiceDownloadRequestRow`: `Certificado de retención <uuid> SERVICIOS
COMERCIALES AMAZON MEXICO ANE140618P37 Amazon.com.mx 1 de septiembre de 2026 30
de septiembre de 2026 pdf xml`. Older rows (2024–2025) are monthly fee invoices
of the previous scheme — ignore them.

**Download:** the `pdf` / `xml` links (`a#download_link_clicked`) submit a form
(hidden inputs `invoiceNumber`, `documentType` RETENCIONES / RETENCIONES_PDF,
`fileExtension`); a real click lands `<uuid>_<n>.xml|pdf` in the download folder.

XML: CFDI `retenciones:Retenciones` v2.0, `Periodo` Ejercicio/MesIni/MesFin,
`Totales` MontoTotOperacion (sales excl. IVA), MontoTotRet, `ImpRetenidos`
(ImpuestoRet 002 = IVA, 001 = ISR), complement `ServiciosPlataformasTecnologicas`
(TotalIVATrasladado, TotalIVARetenido, TotalISRRetenido, NumServ, one
`DetallesDelServicio` per sale). September 2026: operations 426 141.61, IVA
withheld 34 091.32, ISR withheld 10 653.54 (2.5 %).
