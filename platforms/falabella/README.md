# Falabella (Chile)

**Audience:** developers working on this platform, and finance reviewing the
report's rules. Operators use `README.md` at the repository root.

Market code `FA-CL` (core/markets.py). Stores are declared in `config.json`:

```json
"catalog": {"DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"}},
"batches": {"FA-CL": ["DAB_CL-Fala (Coco)", "AIS_CL-Fala"]}
```

| Command | What it does |
|---|---|
| `python -m platforms.falabella.run_batch --month 2026-09 --stores "DAB_CL-Fala (Coco)"` | download the 7 forms of a month for these stores |
| `python -m platforms.falabella.run_batch --market FA-CL --month 2026-09` | every store of the market (`batches["FA-CL"]`) |
| `python -m platforms.falabella.run_batch --collect-only --month 2026-09 --stores …` | fetch reports that were requested but not ready |
| `python -m platforms.falabella.run_downloads --store … --only stock` | one store, some routes |
| `python -m platforms.falabella.run_reports --month 2026-09` | per-store workbooks + `CL_FA_全店汇总_2026-09.xlsx` |

Downloads land in `downloads/<store>/<YYYYMMDD_HHMMSS>_m<YYYYMM>/` with
`run_meta.json` and `REPORT.md`; reports in `downloads/data/reports/financial/`.
Where each form is and how it is fetched: `FORMS.md`.

## When to run

The report is about **last month**. Forms order, billing (Transacciones), stock,
billing2 and invoice2 are complete from the 1st. The second fee-invoice period
(21st of last month to the 20th of this month) appears only after the 20th; until
then the report shows that part of bridge B as "pending" and everything else as
final. Re-running downloads and reports after the 21st completes it; no P&L
number changes, only the check.

## Platform notes

* UI must be **Spanish** (`data-localefull="es-CL"`). The owner's intake
  screenshots were from an English UI and are only a map.
* New pages are micro-frontends in shadow roots; `page_kit.py` walks them.
* Orders: the date filter is in the panel that opens when you click into the
  search box; at most 14 days per export; each export downloads as `export.xlsx`
  and is renamed `Ordenes_<from>_<to>.xlsx` at once. A range without orders
  leaves "Exportar" disabled — recorded as NO_DATA.
* Stock: Chrome may silently refuse the export link after a few downloads from
  `storage.googleapis.com` ("multiple automatic downloads"); the route then
  fetches the same public URL directly (only that bucket) and records
  `stock_fetched_direct` in `run_meta.json`.
* The stock export is a snapshot of the moment of the run, not of month end.

## Accounting rules (for finance to confirm)

These are the rules `report_build.py` implements. They were chosen to be sound
for a Chilean IVA taxpayer and to keep the report a sister of MercadoLibre's;
items marked ⚠ are for the accountant.

1. **Period basis — transaction date.** Each money line of "Transacciones" counts
   in the month of its *Fecha de transacción*. For the sale itself this is about
   the delivery date — when Falabella books and pays it. Fees, refunds and
   promotions carry their own dates. This is also the basis of Falabella's fee
   invoices, so bridge B ties line by line.
   *Order-date* GMV (orders created in the month, from the order list) is shown as a
   KPI; bridge A explains the difference order by order: orders not settled yet,
   and earlier orders settled this month.
2. **Net of IVA.** The P&L uses *Monto (Sin IVA)*. Output IVA (on sales,
   shipping income, promotions) and input IVA (on Falabella's fees) are memo lines,
   not income or cost. Bridge C goes from the contribution to the settled amount
   with IVA. *Differs from the MercadoLibre Chile report*, which is gross because
   MercadoLibre's files do not split IVA; Falabella's do, line by line.
3. **Revenue lines.** Product price (incl. the part paid with CMR points); shipping
   paid by the buyer (and shipping Falabella pays for the seller's own delivery);
   minus returns; plus Falabella-funded promotions and incentives (coupons, CMR
   "Oportunidad Única", sales targets); plus lost/damaged compensation.
4. **Buyer-paid shipping is a pass-through.** It appears as shipping income and,
   for Falabella-delivered orders, as an equal "reversa" charge. Both lines are
   shown so the flow is visible; the net is zero.
5. **Fees** are grouped into fixed rows: commission, logistics co-financing, buyer
   shipping passed to Falabella, free-shipping promotions, reverse logistics,
   Fulfillment (storage, picking, pickups, withdrawals), advertising, subscriptions
   and services, penalties, other corrections. Each transaction type's
   *Reembolso / Corrección / Ajuste* goes to the same row as the charge, so a row is
   a net amount. Every type is registered in `fee_catalog.py` (all 147 of
   Falabella's glossary); an unknown type is an **error** check and its amount is
   kept in a visible "未登记科目 ⚠" row.
6. **Product cost** is not in Falabella's data: yellow input row in ②, unit cost
   column in ⑦.
7. **Cash.** The settled amount with IVA of the month's lines is what Falabella
   pays out through Fpay statements (*Nº estado de cuenta*); lines "No Pagado" are
   paid on the next settlement date. The bank side (Fpay withdrawals) is not
   downloaded yet.
8. **Sales documents.** "Liquidación factura" lists the boletas, facturas and
   credit notes Falabella issued **in the seller's name** — essentially for
   Fulfillment-by-Falabella orders. ⚠ Who issues the documents for the seller's
   own-delivery (FBS) orders, and that output IVA is declared exactly once, is for
   the accountant.
9. ⚠ Input IVA on Falabella's monthly fee invoices (Factura + Nota de Crédito,
   21st–20th) — deductible in full?
10. ⚠ Falabella-funded promotions and compensation are treated as revenue
    adjustments; their invoicing treatment is for the accountant.

Checks (⑩): required files present; only the month's lines used; every type
registered; P&L = sum of lines; IVA = 19% per line; bridge A (warn); bridge B per
closed period (error), open period (warn); unpaid lines (warn).

Golden case: add DAB_CL-Fala (Coco) 2026-09 to `tests/test_golden.py` once
finance has accepted the month (docs/NEW_PLATFORM_SOP.md §5).
