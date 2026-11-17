# MercadoLibre — platform notes

**Audience:** developers changing the MercadoLibre downloads or reports. The
form-by-form detail is in `FORMS.md` (same folder); operators read the root
`README.md`.

## Countries (sites)

Per-country facts live in `meli_site.py`; a store's country comes from
`config.json` → `stores.regions` (default MX).

| | Mexico `mx` | Chile `cl` | Peru `pe` |
|---|---|---|---|
| Ziniao platform name | MercadoLibre-墨西哥-本土 | MercadoLibre-智利-本土 | MercadoLibre前台-秘鲁 |
| seller centre | vendedores.mercadolibre.com.mx | vendedores.mercadolibre.cl | vendedores.mercadolibre.com.pe |
| MercadoPago | mercadopago.com.mx | mercadopago.cl | mercadopago.com.pe |
| UI language required | es_MX | es_CL | es_PE |
| currency | MXN, 2 decimals | CLP, **no decimals** | PEN, 2 decimals |
| timezone | America/Mexico_City | America/Santiago | America/Lima |
| sales sheet | "Ventas MX", columns "(MXN)" | "Ventas CL", "(CLP)" | "Ventas PE", "(PEN)" |
| monthly bill | calendar month | **per-store cut day** (e.g. 8th→7th, 24th→23rd); charges are assigned by charge date | calendar month |
| Full warehouse (stock reports) | yes | yes | **no** — stock pages are 404; not downloaded, not expected |
| tax | IVA 16%; platform withholds IVA/ISR | IVA 19%; nothing withheld | IGV 18%; nothing withheld; fee invoices are boletas (`B001-…`) |

## Quirks found in production (keep them in mind)

- **Peru spells September "Setiembre"** in the seller centre (bill cards, detail
  page). Month regexes use `Sep?tiembre`. MercadoPago's date picker still says
  "septiembre", so MercadoPago month names are not switched per site.
- **Peru's sales-export notification clears** after a while: the export is
  collected right after requesting it (`ventas_collect_now` in `meli_site.py`).
- **Chile's returns menu** says "Reporte de devoluciones al centro de
  distribución" (Mexico: "… a la bodega"); both are accepted.
- **New accounts**: MercadoPago greys out dates before the account existed; the
  start date moves to the first selectable day.
- **Buyer-paid shipping** (Chile, Peru) and the Flex delivery bonus (Peru) are
  separate MercadoPago payments, so the per-order MercadoPago comparison shows
  explained differences.
- **Platform says the sales export is incomplete** ("No pudimos mostrar algunos
  datos"): the report is still built, with a warning.
- **A downloaded .xlsx/.csv that is not one** (the platform served an error page
  under the report's name; EWTTO_SM 2026-10-30, five stock-type files): both
  bases treat the file as absent and record a failed check `load:<file>` in ⑩
  (order basis: `S["load_errors"]`, `report/checks.load_error_checks`; warn level
  for stock-type files and Cargos Full, which do not change ②; error for Pagos
  and MercadoPago). Only Ventas and Facturación still fail loudly — the report
  cannot exist without them.
- The machine clock is China time; "today" for the platform is computed in the
  site's timezone (`core/site_time.py`).

## Reventa (MercadoLibre buys from the seller and resells)

MercadoLibre Mexico invites listings into Reventa: at the buyer's click the store
sells the item to MercadoLibre ("vendido por Mercado Libre"), invoices MercadoLibre
(CFDI G01) and is paid a net amount. Measured on EWTTO_SM, whose whole store
switched on 2026-10-26 about 12:10 (download of 2026-11-11; no other store checked
yet). The sales export says which sale is which (`Modelo de venta`), so a store, a
listing or a single day can be mixed.

| | direct sale | Reventa sale |
|---|---|---|
| `Ingresos por productos` | the buyer's price | the net amount MercadoLibre pays (EWTTO_SM: 74 % of the suggested price) |
| `Cargo por venta e impuestos`, `Costos de envío` | commission + withheld IVA/ISR, shipping | 0 |
| unit-price column | the listing price | the seller's suggested price |
| monthly bill | `Cargo por venta`, `Cargo por envíos` rows | **no row at all** |
| withholding (Mexico) | IVA/ISR withheld by MercadoPago | none |
| ads, Full storage, Full collection | billed | billed as before |
| cancellation / return | commission reversal in the bill | `Anulaciones y reembolsos` = minus the amount, Total 0 |

How the reports treat it (a store-month without Reventa sales gets the same
workbook, cell for cell, as before — the golden cases prove it):

- **Bill basis** (`bill_report/model._reventa`): the bill cannot carry these sales,
  so they come from the sales export and enter the same machinery as a billed sale —
  amount = the net amount, date = the sale date (there is no charge date), SKU from
  the export row, reversal = the row's `Anulaciones y reembolsos` as a share. They
  are in ② 商品销售额 and the reversal rows, in ⑦ per SKU (units count for cost), in
  ⑧ by week, and in the base that store-level charges are allocated on. No
  commission, shipping, buyer shipping or withholding. ② 备查项 splits Reventa from
  direct sales and shows the suggested-price total; ⑦ gets 表三 (direct next to
  Reventa per SKU); ① 五 states the policy; ⑩ a check `reventa` and three open
  questions.
- **Order basis**: its P&L already sums the export, so the totals and the bridges
  hold. Added: the withholding rate is measured on direct sales only
  (`m["ing_direct"]`; with Reventa in the denominator the `tax_rate` check failed),
  the same memo split in ②, the comparison at the end of ⑦, a line on ①, an item on
  ⑩, a check `reventa`.
- **Known limits.** The export does not date a refund: a Reventa refund is in the
  month of its sale as of the download; ② 备查项 (bill basis) shows last month's
  Reventa refunds as of this download so a later refund is seen. The amount is
  treated as including IVA like every other line — to confirm against one CFDI
  issued to MercadoLibre. How and when MercadoLibre pays was not verified: the
  MercadoPago reports of a Reventa store have not been seen (the VM login has no
  reports permission).
- `report/reventa.py` and `bill_report/reventa.py` are the same file twice (the
  bases never import each other; `tests/test_ml_reventa.py` keeps them identical).

## Forms not downloaded on purpose

- MercadoPago account statement ("Estados de saldos y movimientos"): the
  "Generar" button is often not found (always, so far, in Chile and Peru); the
  report does not use it.
- `BILL_MP_V3` (fifth billing file): asks for a format in a dialog the script
  does not answer; the report does not use it.
- Both are recorded with a cause but do not count as "report missing data"
  (`NOT_IN_REPORT` in `core/download_diag.py`).

## Accounting rules

The rules, the ten sheets and the per-country differences are documented for
operators and finance in the root `README.md` §3.6 (countries) and §4 (sheets).
Open accountant questions per country are listed in the run reports' ⑩ sheet.

**Product cost** is not in MercadoLibre's data. Without a cost table: yellow input
row in ②, yellow unit-cost column M in ⑦ 表五. With the shared SKU cost table
(`core/sku_costs.py`, see `docs/REPORT_FAMILY.md` §7), platform code `ML`,
country `MX` / `CL` / `PE`, currency MXN / CLP / PEN:

- **Cost units = ⑦ 表五's 销售件数 (column E)** — the units of the month's orders
  (order date, the same Ventas rows as ②'s revenue lines), minus orders cancelled
  before shipping (表二: nothing left the warehouse) and minus returns that ML's
  review says are back on sale (表三: the goods are back in stock). Returns judged
  not resellable or not reviewed yet, mediations, claims, exchanges and
  undelivered orders (表四) keep their cost — the goods are lost or their fate is
  open, so the conservative choice is to count them; the refund of their revenue
  is already in ②. Why this basis: it is the one ⑦ already used for
  `采购成本合计 = E × M`, so ⑦'s cost and ②'s cost are the same number, and it is
  revenue's basis (order date, same month). A cancelled-before-shipping order
  never has a cost.
- `SC.Item(sku, name, units=E, returned=0, sales=表五 GMV)` per 表五 SKU
  (`report/sheet_sku.py: sku_items`); `returned` is 0 because the excluded units
  are already out of E.
- **Unit cost** = landed cost per unit as finance enters it (purchase + freight
  + duties), **net of recoverable IVA / IGV**. ML's revenue lines include VAT
  (Mexico, Chile, Peru), the cost does not; the policy line in ① says so.
- ② 商品采购成本 = the covered SKUs' cost (still a yellow cell; note: file and
  coverage); ⑦ 表五 M = the unit cost (yellow 0 where uncovered), AE 成本来源 =
  table row used or 未录入成本; ① 一 gets the file, 六 the cost and profit, a new
  七、商品成本（会计处理）the policy; ⑩ gets the cost-table checks and, while SKUs
  are uncovered, the to-do "补录商品成本表". ⑨ values stock at the same unit cost.
- `report_build.sku_list` and `run_reports.latest_folder` let
  `tools/sku_cost_template --market ML-MX` list a month's SKUs.

## Tests

`tests/test_cl_*.py`, `tests/test_pe_*.py` (synthetic data) and
`tests/test_golden.py` (golden: real Mexican, Chilean and Peruvian store-months
rebuilt cell by cell against the accepted workbooks; data in `~/report-golden/`,
outside git).
