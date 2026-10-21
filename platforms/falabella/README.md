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

The report is about **last month**, and every number in it is final from the 1st of
the next month: all seven forms for the month exist then, and the P&L comes from
"Transacciones", which already holds every sale, fee and refund of the month with
its own date and IVA.

The one thing that comes later is Falabella's **fee invoice** for the period that
runs past the month end (21st of last month to the 20th of this month). It is only
a cross-check of fees that are already in the P&L line by line (③ bridge B), and
it is the tax document for the input IVA of those days — which belongs to the
following month's IVA return anyway, because input IVA is declared by invoice
date. So nothing in the month's report is missing before then; bridge B shows those
days as "pending". The next month's run downloads that invoice as part of its own
month (it covers the 1st–20th of that month), so no manual re-run is needed;
re-running the old month after the 21st only turns its "pending" into "checked".

## Platform notes

* UI must be **Spanish** (`data-localefull="es-CL"`). If a store is found in
  another language (someone switched it), the run switches it back to Español
  through the footer language box and notes it in `REPORT.md`; only when that fails
  does the store stop with LANGUAGE_WRONG. The owner's intake screenshots were from
  an English UI and are only a map.
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

## Accounting rules

Decided for a Chilean IVA taxpayer reporting under IFRS-style principles, and kept
a sister of MercadoLibre's report. The workbook states the same rules on ① (七、会计
处理) and the points that matter for the tax return (八、报税时要注意).

1. **Revenue recognition — delivery (IFRS 15).** A sale is revenue when Falabella
   records "Pago por precio del producto" for it, i.e. at delivery, when control
   passes to the buyer. Every money line counts in the month of its *Fecha de
   transacción*. Orders placed but not delivered in the month are not revenue;
   order-date GMV is a KPI, reconciled order by order in ③ bridge A.
2. **Net of IVA.** The seller is an IVA taxpayer: output IVA and input IVA are
   neither income nor cost. The P&L uses *Monto (Sin IVA)*; both IVA amounts are
   memo lines and ③ bridge C walks from the contribution to the cash with IVA.
   (MercadoLibre Chile's report is gross only because MercadoLibre's files do not
   split IVA; Falabella's do, per line.)
   **IVA reference (added 2026-10-21, adds cells only — no P&L row or column moves):**
   below ②'s memo, "IVA 明细 · 按损益表行" lists the IVA of every income and fee row
   (transaction date, the same lines as the P&L; buyer shipping as its two memo legs),
   their totals (= the memo's 销项 / 进项 IVA) and 应缴 IVA 净额 = 销项 − 进项, so that
   contribution + 应缴 IVA 净额 = the settlement with IVA. "IVA 明细 · 按单据日期" gives
   the F29 reference by document date: input IVA of the fee invoice whose period ends in
   the month (all its lines, from InvoiceReport; 0 with a note when not downloaded),
   output IVA of the Liquidación documents (estimated as total × 19/119) and of FBS
   sales and returns (the seller's own documents). ⑦ gets 销项 / 进项 IVA columns and ⑧
   销项 / 进项 / 应缴 IVA columns, at the right end.
3. **Shipping — principal vs agent.** For orders Falabella delivers, the shipping the
   buyer pays is passed to Falabella unchanged ("Pago de envío comprador" and the
   equal "Reversa"): the seller is an agent, so it is **not revenue** (net
   presentation; both legs are memo lines). Shipping the seller earns — buyer-paid
   shipping on own-delivery orders, and what Falabella pays for the seller's own
   delivery — is revenue ("运费收入（净额）").
4. **Promotions.** Falabella-funded coupons, CMR "Oportunidad Única" and sales
   incentives are part of the consideration for the sale: revenue adjustments in net
   sales. Seller-funded free-shipping promotions are a selling cost (shipping row).
5. **Returns.** Contra-revenue ("退货退款"); the commission Falabella refunds on a
   return reduces the commission row; reverse logistics is a logistics cost.
6. **Lost/damaged compensation** compensates an inventory loss, it is not a sale:
   "其他收入", below the platform fees.
7. **Platform fees** count on Falabella's charge date (the basis of its fee invoice),
   in fixed rows: commission, logistics co-financing, free-shipping promotions,
   reverse logistics, Fulfillment (storage, picking, pickups, withdrawals), ads,
   subscriptions and services, penalties, other corrections. Each type's *Reembolso /
   Corrección / Ajuste* nets in the same row. All 147 types of Falabella's glossary
   are registered (`fee_catalog.py`); an unknown one is an **error** check, its
   amount kept in "未登记科目 ⚠".
8. **Product cost** is not in Falabella's data. Without a cost table: yellow input
   row in ②, unit cost column in ⑦. With the shared SKU cost table
   (`core/sku_costs.py`, see `docs/REPORT_FAMILY.md` §7): cost of goods = unit
   cost × (units settled in the month − units returned in the month), i.e. the
   count of "Pago por precio del producto" lines minus the count of "Descuento por
   devolución de producto" lines of the SKU, on the same transaction-date basis as
   the revenue lines; a returned unit is taken as back in stock. Unit cost = the
   landed cost per unit as finance enters it (purchase + freight + duties, net of
   IVA, in CLP — no FX conversion). ② 商品采购成本 holds only the covered SKUs'
   cost (still a yellow cell, its note gives the coverage); ⑦ shows each SKU's cost
   and margin, "未录入成本" where the table has none. The stock value on ⑨ stays at
   selling price — a size indicator, not the accounting value of inventory; valuing
   stock at cost is a later step.
9. **Cash.** The month's lines with IVA are what Falabella settles through Fpay
   statements (*Nº estado de cuenta*); "No Pagado" lines are paid on the next
   settlement date.

Points for the IVA return (they do not change the report):

* Output-IVA documents for Fulfillment-by-Falabella orders are issued by Falabella in
  the seller's name ("Liquidación factura", ③ bridge D) — declare those, do not
  issue them again. Own-delivery (FBS) orders need the seller's own documents
  (amount on ③ bridge D).
* Input IVA on Falabella's fee invoices is declared by invoice date: the invoice
  for 21st–20th is issued after the 20th, so fees of the 21st–end of a month go
  into the next month's IVA return.
* Promotion and compensation documents are Falabella's (Nota de crédito / Nota de
  débito promociones).

Checks (⑩): required files present; only the month's lines used; every type
registered; P&L = sum of lines; IVA = 19% per line; bridge A (warn); bridge B per
closed period (error), open period (warn); unpaid lines (warn). With a SKU cost
table: file readable, rows valid, no duplicate keys, currency = CLP (error);
coverage of the SKUs sold (warn, uncovered SKUs listed by sales).

Golden case: add DAB_CL-Fala (Coco) 2026-09 to `tests/test_golden.py` once
finance has accepted the month (docs/NEW_PLATFORM_SOP.md §5).
