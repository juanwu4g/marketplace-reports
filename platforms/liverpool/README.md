# Liverpool (墨西哥)

**Audience:** developers working on this platform, and finance reviewing the
report's rules. Operators use `README.md` at the repository root.

Market code `LV-MX` (core/markets.py). A Liverpool seller needs **two Ziniao
stores**: the supplier-portal user (the store the report is named after) and a
marketplace user (orders, shipping charges). Declare both in `config.json`:

```json
"allowed": ["Liverpool Report", "Liverpool Order 子账号(Nikki)"],
"catalog": {"Liverpool Report": {"platform": "LV", "country": "MX",
                                 "marketplace_store": "Liverpool Order 子账号(Nikki)"}},
"batches": {"LV-MX": ["Liverpool Report"]}
```

The marketplace user must be on the allowlist; it is never run as a store of its
own (it has no catalog entry). In Ziniao, the supplier-portal store needs the
portal's user and password saved, so that Ziniao fills the sign-in form. The
script never types a password (see FORMS.md "Signing in"). Ziniao must also allow
`itsproves.liverpool.com.mx` (whitelisted 2026-10-26).

| Command | What it does |
|---|---|
| `python -m platforms.liverpool.run_batch --month YYYY-MM --stores "Liverpool Report"` | download the month's forms (both Ziniao stores) |
| `python -m platforms.liverpool.run_batch --market LV-MX --month YYYY-MM` | every store of the market |
| `python -m platforms.liverpool.run_downloads --store … --only ledger sales` | one store, some routes |
| `python -m platforms.liverpool.run_reports --month YYYY-MM` | per-store workbooks + `MX_LV_全店汇总_YYYY-MM.xlsx` |

Downloads land in the store's run folder (core/layout.py), together with
`run_meta.json` and `REPORT.md`. Where each form is and how it is fetched:
`FORMS.md`.

Status (2026-10-26):
- **Report:** implemented and checked against the real September 2026 files of
  `Liverpool Report`. Every bridge ties: orders → booked sales, commission per
  sale, shipping list → ledger, P&L → ledger. Tests on synthetic files are in
  `tests/test_lv_report.py`.
- **Downloads:** implemented from the probed routes (`tests/test_lv_downloads.py`
  covers the two-store runner).

## What the report reads

| Form | File | Used for |
|---|---|---|
| ledger | `LV_mayor_<YYYYMM>.xlsx` | **every number** of ② (both accounts, posted in the month) |
| sales | `LV_ventas_<YYYYMM>.xlsx` | ledger document → order → SKU; ③ bridges A, B |
| orders | `LV_pedidos_<date>.xlsx` | seller SKU and product name per order; order-date KPIs; ③ A; the "not delivered" memo |
| logistics | `LV_costos_logisticos_<YYYYMM>.xlsx` | shipping per order (⑦); ③ D |
| balance | `LV_balance_<YYYYMM>.xlsx` | extra shipping per order (⑦) |
| payments | `LV_pagos_<YYYYMM>.xlsx` | ③ C reference line |
| invoices | `Notas_Cargo_<YYYYMM>/` | tax documents (not read for figures) |

A file missing in the folder is taken from another run of the same month, as on
Falabella and Amazon. Missing orders / logistics / balance lists leave ② unchanged
(core/grading NON_BLOCKING_FORMS): ⑦ then shows Liverpool's article number and
allocates shipping by sales.

## Accounting rules

Decided for a Mexican IVA taxpayer, and kept a sister of MercadoLibre Mexico's
official (bill) basis. The workbook states them on ① (六、会计处理) and the
points for the tax return (七).

1. **Basis: Liverpool's ledger.** Every line of the supplier ledger posted in the
   month counts in the month (Fecha de contabilización), and only those. This is
   the same idea as MercadoLibre's bill basis: the platform's own record decides.
   - **Measured on 2026-09:** Liverpool posts a sale on the day the buyer orders
     and pays (122 of 122 sales on their order day, about 6 days before delivery).
     A cancellation or return is a reversal document on its own day.
   - **Delivery timing:** sales posted in the month and not delivered by its end
     are a memo line on ② ("本月入账、月末尚未送达的销售"). That is the timing
     difference to recognising revenue at delivery (IFRS 15).
   - **Order-date figures:** order-date GMV and order counts are a KPI, reconciled
     order by order in ③ A.
2. **With IVA**, as Liverpool books and pays: the same basis as MercadoLibre
   Mexico's official report for the same company. The IVA inside the figures is a
   memo: output = net sales × 16/116, input = fees × 16/116.
3. **Liverpool-funded promotions** (the buyer pays less, Liverpool pays the seller
   the difference with its own credit document) are part of the consideration:
   buyer-paid price + promotion credit = list price, both revenue. Liverpool's
   commission is on the list price.
4. **Cancellations and returns** are contra-revenue ("减：取消与退货退款"). The
   commission Liverpool gives back reduces the commission row; the withholding it
   gives back reduces the withholding row.
5. **Platform fees** in fixed rows:
   - commission (net of returns);
   - non-Liverpool-card surcharge;
   - shipping;
   - extra shipping;
   - logistics adjustments of earlier months;
   - interest-free instalment (MSI) cost;
   - subscription (Cuota digital);
   - other;
   - "未登记科目 ⚠".

   Every kind of ledger line is registered in `catalog.py`; an unknown one is an
   **error** check and its amount goes to the unregistered row. The fees sit on a
   second account and are netted with the payments.
6. **Withheld taxes**: IVA 8 % and ISR 2.5 % of the price net of IVA, given back
   on returns. They are a prepayment of the seller's taxes, not a cost: a row
   below the profit, with a sensitivity row "if not deductible". This is exactly
   MercadoLibre Mexico's "平台代扣代缴税金".
7. **Product cost** is not in Liverpool's data:
   - **With the shared SKU cost table:** cost = unit cost × 1.16 (the table is
     net of IVA, the report is with IVA) × (units posted − units returned in the
     month). The SKUs are the seller SKUs of the order list.
   - **Without the table:** yellow input cells on ② and ⑦.
   - **No platform stock:** no Liverpool Fulfillment (FBL) is used, so ⑨ is not
     applicable.
8. **Attribution to SKUs (⑦)**:
   - **Sale lines:** sale, promotion, return, commission, surcharge and
     withholding follow their document's orders (Ventas Marketplace); a document
     with several orders is split by price.
   - **Shipping:** follows the shipping list's orders (by Doc contable).
   - **Extra shipping:** follows Balance de cuenta's orders.
   - **Everything else** (subscription, adjustments, shipping documents not in
     the list) is allocated by each SKU's net sales, negative counted as 0.
9. **Cash**: the payments (ZP) net of the fees settled with them. A payment the
   bank sent back is reversed in the ledger and nets to 0; it is a warn check
   (3 in September 2026, before the payment of 28.09 went through).

Checks (⑩):
- the ledger and Ventas Marketplace present (error);
- the marketplace lists present (warn);
- only the month's lines used;
- every line registered (error);
- P&L = ledger (error);
- every sale document linked to an order (warn);
- bridges A, B, D (warn);
- payments sent back (warn);
- with a SKU cost table, its coverage.
