# Experiment: MercadoLibre on the invoice basis

**This is an experiment, not part of the monthly run.** Nothing in `core/`,
`platforms/`, the runner, the window or the scheduled tasks imports it, and it
writes its own, differently named workbooks. The monthly report
(`<store>_<YYYY-MM>.xlsx`) is unchanged; deleting this folder changes nothing else.

> 给老板的一句话：这是一份**实验报表**——把损益的依据从销售报表（Ventas）换成当月账单
> （Facturación + Notas de crédito），看看能做到多好、和现行报表差多少、哪个更稳定。
> 每月的正式报表不受影响。

## The question

The owner: *"The order form (Ventas) is changing constantly due to changes and
returns — how different would it be if the basis were the month's invoice
(`Reporte_Facturacion_MercadoLibre_<Mon><YYYY>.xlsx` +
`Reporte_Notas_Credito_MercadoLibre_<Mon><YYYY>.xlsx`)?"*

## Commands

```
# one store-month: the invoice-basis P&L, the comparison, the details, and ⑥ stability
.venv/bin/python -m experiments.ml_invoice_basis --store <name> --month 2026-09 \
    [--folder DIR] [--base DIR] [--out-dir DIR] [--no-recalc] [--no-stability] \
    [--threshold-pct 1.0] [--threshold-abs 1000]

# several downloads of the same month: how much each basis moved, and why
.venv/bin/python -m experiments.ml_invoice_basis.stability --store <name> --month 2026-09 \
    [--folders DIR DIR ...] [--base DIR] [--out-dir DIR] [--no-recalc]
```

* Default folder: the newest `downloads/<store>/*_m<YYYYMM>` that holds the month's
  Facturación. Default stability folders: every `*_m<YYYYMM>` of the store, oldest first.
* Output: `downloads/data/reports/experiments/` (git-ignored) —
  `<store>_<YYYY-MM>_账单口径实验.xlsx` and `<store>_<YYYY-MM>_口径稳定性.xlsx`.
* `--threshold-pct` / `--threshold-abs`: materiality of ② (a difference is "重要" when it is
  ≥ this % of the current basis' net revenue, or ≥ this amount in the store's currency);
  both are also yellow input cells on ②.
* Cost of goods appears only when `core/sku_costs.load_configured()` finds a SKU cost
  table (`config.json` → `reports.sku_costs`); otherwise ③ says "未配置 SKU 成本表".
* The site (MX / CL / PE) comes from each folder's `run_meta.json` (none = MX); the
  previous site is restored afterwards. `config.json` is never read. Nothing opens a
  browser or a store.
* Exit codes: 0 written; 1 written but a formula error was found after the
  LibreOffice recalculation; 2 a problem the person must fix (Chinese message).

## Workbooks

`<store>_<YYYY-MM>_账单口径实验.xlsx`

| Sheet | Content |
|---|---|
| ① 说明 | what this is, **the periods**, the files (invoices, Pagos_Facturas, MercadoPago), **the VAT proof**, what the invoice can / cannot give, **the accounting policies**, methods, key numbers |
| ② 差异摘要（给财务） | for finance: a walk from the current contribution to the invoice basis (calendar cut): every differing line, the adjustments to actual, the VAT presentation; amount, % of net revenue, "重要" flag (thresholds in yellow cells), reason category (时间差 / 账单里没有 / 估算与实际 / 取数来源不同 / 含税/不含税列报); three checks = 0 |
| ③ 账单口径损益 | invoice-basis P&L: invoice period and calendar cut, each **gross and net of VAT**; residual check; "调整为实际（MercadoPago）与应计" (estimate added back, actual refunds, coupons, accrual → adjusted contribution; estimate error and dispute charges as memo rows); cost of goods (or "未配置 SKU 成本表"); VAT memo (output / input / payable); "发票口径无法得到" with the withholding certificate check |
| ④ 与现行订单口径对比 | line by line (gross, reversals estimated — the earlier comparison, unchanged): current / invoice period / calendar cut / differences / reason / explanation; contribution check; order ↔ billed-sale reconciliation and an exact GMV-difference split |
| ⑤ 费用明细 | every FEE_CATALOG line: charged, reversed, net; unregistered Detalle values in red rows |
| ⑥ 销售与冲销明细 | reversals with link method and ratio; sales billed but not orders of the month; orders of the month not billed in the period; amount mismatches; every billed sale |
| ⑦ 退款优惠券与现金 | MercadoPago files present / used; reversals: estimate vs actual per reading and per sale (status, error); coupons and their check against Ventas; disputes; withheld tax; **cash bridge** to the money released in the month; where the money went |
| ⑧ 三种日期与应计 | accrual rows (later invoice, open/closed); revenue by month under order / billing / delivery date; per sale the three dates |
| ⑨ 口径稳定性（多次下载） | the stability table when sibling downloads of the month exist (now with the actual / net-of-VAT lines) |

`<store>_<YYYY-MM>_口径稳定性.xlsx`: ① 说明 (conclusions) ② 稳定性总表 ③ 订单变化明细
(Ventas, order by order) ④ 账单行变化明细 (invoice rows, charge by charge).

## Method

**Two readings of "the month"** (both from invoice rows only):

* **Invoice period (账单期)** — the documents named for the month. Mexico and Peru
  bill by calendar month. Chile bills from the store's cut-off day (inferred per
  store by the report's own `infer_cycle_day`; DAB_CL: the 8th): "Sept2026" covers
  **Aug 8 – Sep 7**.
* **Calendar cut (自然月切)** — rows of all downloaded invoices whose charge date
  (`Fecha del cargo`) is in the month. Chile September = Sep 1–7 from "Sept" + Sep 8–30
  from "Oct". A missing or still-open invoice is named on ① and ②.

**Revenue.** One billed sale per `Número de venta` of a commission charge
(`Cargo por venta`; an affiliate charge on the same sale adds commission only):
`Total de la venta`, `Cantidad vendida`, `Precio unitario`, `Fecha de venta`, listing.

**Reversals.** `Anulación del cargo por venta` (in Facturación or Notas de crédito)
undoes a share of a sale: *share = reversed commission ÷ original commission*,
*reversed sales = share × sale total*. MercadoLibre returns the commission in the same
proportion as the refund (the monthly report's credit-note audit proved it against
MercadoPago cash). The original commission is found, in order:

1. `Cargo que bonifica` / `Cargo que anula` → `Número del cargo` (any downloaded invoice);
2. the same `Número de venta` in any downloaded invoice;
3. the original charge is in an invoice not downloaded: recompute it from the
   listing's fee terms (`Porcentaje por categoría` × total + `Costo por unidad vendida`
   × units) seen on its other sales — a term that reproduces the reversed amount exactly
   means a full reversal, otherwise the term of the nearest sale date. (A median rate
   was tried first and rejected: a listing's rate changed from 14.5 % to 11 % in
   October, so the median moved when a later invoice was downloaded and the
   "closed" Sept invoice basis changed by 892.85 between two downloads.)
4. otherwise a full reversal.

**Buyer-paid shipping is in the invoice** (contrary to the first assumption):
`Envío a cargo del cliente` on the `Cargo por envíos` row equals Ventas' `Ingresos por
envío` order by order (DAB_CL 592/592, BOCINA_TA02 13/13). A reversed shipping charge
returns it in the same proportion.

**Fees.** By FEE_CATALOG line (same Chinese names as the monthly report), charged −
reversed (credit notes included) = net; any Detalle not in FEE_CATALOG gets its own red
row and is never merged. Σ lines = Σ invoice rows is checked on ② and ④.

**Not in the invoice** (shown as "发票口径无法得到", with the order-basis value as a grey
reference, never added into invoice totals): withheld IVA/ISR (Mexico), seller coupons,
the net refund actually paid to buyers, Peru's Flex shipping bonus, Ventas `Total` /
MercadoPago cash, cost of goods.

**Order basis** = the monthly report's own `derive()` + `build_context()` contribution,
imported read-only; ③'s check row proves the line-by-line sum equals it.

## The accountant's additions (2026-10-20, second round)

All from files already downloaded; each number says its source, and a missing file is
written per line ("本次下载没有 MercadoPago 文件"), never a failure.

**VAT inside `Valor del cargo` — proved, not assumed** (`vat.py`, ① 四). Four tests:
(1) Pagos_Facturas "Detalle de Pagos del mes": the part of each payment applied to a
charge = that charge's `Valor del cargo`; (2) each credit note's `Importe total` (a fiscal
document) = −Σ `Valor del cargo` of the rows with its `N° de factura fiscal`; (3) the
commission row = `Costo por categoría` + `Costo por unidad vendida` − discount, nothing on
top; (4) MercadoPago FEE_AMOUNT of a payment = Σ "Si" charges of it. Real data (2026-09):

| | (1) payments | (2) credit notes | (3) commission rows | (4) MercadoPago fee |
|---|---|---|---|---|
| BOCINA_TA02 | 219/219 | 8/8 | 436/436 | 376/404 |
| DAB_CL | 1,283/1,283 | 1/1 | 3,866/3,866 | 3,622/3,785 |
| EWTTO_SM | 489/489 | 29/29 | 971/971 | (no MercadoPago) |
| TANKE_EE | 401/401 | 38/38 | 1,240/1,240 | 1,101/1,162 |

Were the charges net of VAT, cash and credit notes would be 1.16 / 1.19 × the charges;
they are 1 ×. Since MercadoLibre invoices its services with VAT, the VAT is inside:
amount × r/(1+r). Assumption (stated on ①): every charge carries the standard rate; the
per-invoice tax is in the CFDI / DTE XML, not downloaded. Input VAT is creditable in MX /
CL; in PE the fee documents are boletas (no IGV credit), so PE fees stay gross. On
TANKE_EE 392 commission rows have `Costo por categoría` ≈ 0.95 × rate × price: a reduced
rate, not a tax. (The current report: all gross; "账单口径纯佣金" means "without the
withheld tax", not "without VAT"; `iva_in_gmv` / `net_revenue_exvat` only on bridge F.)

**MercadoPago** (`mercadopago.py`). The invoice's `Pago` = settlement_v2 `SOURCE_ID` =
Cobros `operation_id`. Refund = −TRANSACTION_AMOUNT of `REFUND`, plus `DISPUTE` rows with
money (mediations / chargebacks the buyer won). Seller coupon = TRANSACTION + FEE + TAXES −
REAL of the `SETTLEMENT` row — the coupon the monthly report derives from Ventas, order by
order (TANKE_EE 836/840 orders, BOCINA 197/197, DAB_CL 2,436/2,437); Cobros' `coupon_fee`
is not it. A zero-amount `DISPUTE` with a fee is the invoice's own return fee (`Cargo por
devolución`, same payment and amount: DAB_CL 13/14, BOCINA 4/4, TANKE 2/2) — already in the
fees, never added. Without settlement_v2, Cobros' `amount_refunded` gives the refund
(= settlement on 876/878 TANKE payments); coupons are then "无法得到" (a Cobros-based
coupon matched only 451/756).

**Policies** (① 七, decided as an accountant would): revenue at billing date (delivery
date shown as the IFRS-15 comparison on ⑧); refunds in the reading of their reversal, at
the actual amount where MercadoPago has it (allocated by reversed commission when a sale is
reversed in several invoices; unpaid cancelled sales reversed in full; otherwise the
estimate, counted); coupons as a sales discount, given back with refunds; VAT out of both
revenue and fees (gross columns kept); accrual of order-level charges of the reading's
sales already billed in a later downloaded invoice (not-yet-billed ones are unknown, not
estimated); withholding is a tax credit, its source document CFDI de retenciones is not
downloaded yet (MercadoPago TAXES_AMOUNT shown for reference only); cost of goods only with
a cost table (units = billed − fully reversed); materiality 1 % / 1,000.

**Cash bridge** (⑦ F, calendar month): adjusted contribution + bill-level fees paid apart
("No") + accrual − withheld tax = expected cash of the reading's sales; vs their actual
REAL_AMOUNT (all dates); the difference split payment by payment into payment ≠ price +
buyer shipping, earlier-billed reversed sales, refunds not yet reversed in the invoice,
fees, other order-level rows, coupons — the remainder is shown (0 up to rounding); then to
cash by transaction date (check = 0) and by release date (check = 0), compared with
Liberaciones for the same payments (difference shown, not forced).

**Order ↔ billed sale.** A Ventas multi-product pack is a parent row (`Paquete de N
productos`, with the money) plus N item rows; the invoice bills each item and names the
pack in `Número de paquete`. Billed sales are matched on the pack first, then on their own
id; the GMV difference splits exactly into "billed, not an order of the month" +
"order of the month, not billed in the reading" + "amount differences" (③, residual 0).

## Findings on real data (2026-09, downloads of Oct 13/14 and Oct 20)

| | BOCINA_TA02 (MX) | DAB_CL (CL) |
|---|---|---|
| order basis GMV / contribution | 239,931.96 / 158,981.41 | 47,978,708 / 33,065,458 |
| invoice period GMV / contribution | 247,076.15 / 157,643.26 (Sep 1–30) | 13,389,304 / 10,302,289 (Aug 8 – Sep 7) |
| calendar cut GMV / contribution | same as period | 48,291,399 / 33,707,479 |
| order basis, Oct 13/14 → Oct 20 | contribution −56.46 (affiliate commission of Sept orders charged in the still-open Oct invoice); Ventas unchanged | GMV −3,061,085 and contribution −1,893,241: 97 September orders missing from the Oct 20 Ventas export ("No pudimos mostrar algunos datos"), 1 status change, 1 amount change |
| invoice basis, Oct 13/14 → Oct 20 | identical, row by row (672 rows) | identical, row by row (1,011 period rows, 6,015 calendar rows) |

* The closed invoices did not change at all between downloads; the open Oct invoice
  grew (Mexico) and its rows changed `Estado` / channel when later reversed — so the
  invoice basis is stable **once the invoices of the month are closed** (Mexico: early
  next month; Chile: after the 7th of next month for the calendar cut).
* Mexico: the difference to the order basis is mostly sales Ventas never lists
  (9 orders cancelled right after purchase: billed and reversed) and the reading of
  refunds (gross reversed sales vs. Ventas' net refund).
* Chile: the invoice *period* is a different month (Aug 8 – Sep 7); the *calendar cut*
  is within 0.7 % of the order basis on GMV and 1.9 % on contribution.
* The invoice can also check Ventas' completeness: on DAB_CL every September order of
  the Oct 14 Ventas is billed, and the 97 orders absent on Oct 20 are all still in the
  (unchanged) invoice.

## Second-round results (2026-09, golden folders; TANKE_EE / EWTTO_SM Oct 14)

| | BOCINA_TA02 (MX) | DAB_CL (CL, calendar cut) | EWTTO_SM (MX) | TANKE_EE (MX) |
|---|---|---|---|---|
| contribution: current / invoice est. | 158,981.41 / 157,643.26 | 33,065,458 / 33,707,479 | 674,265.15 / 676,387.94 | 383,876.40 / 394,523.32 |
| adjusted (gross) / net of VAT | 157,643.22 / 135,899.33 | 33,669,869 / 28,294,007 | 675,336.65 / 582,186.77 | 386,650.28 / 333,319.21 |
| reversals: estimate / actual (known) | 17,277.67 / 17,277.71 (18/18) | 1,898,926 / 1,919,080 (82/82) | 118,139.42 / — (no MercadoPago) | 60,776.72 / 58,210.49 (71/71) |
| estimate error (est − actual) | −0.04 | −20,155 | — | +2,566.23 |
| coupons (MercadoPago, net) | 0 | 0 | 无法得到 | 10,133.86 |
| accrual (later invoice) | 0 (Oct 20 download: 56.46) | 17,455 | 1,051.29 | 305.41 |
| cash: expected / actual / difference | 162,272.79 / 161,242.32 / −1,030.47 | 36,426,619 / 35,140,888 / −1,285,730 | — | 391,188.49 / 384,129.68 / −7,058.81 |
| released in the month vs Liberaciones | 158,359.02 / 158,486.32 | 31,226,340 / 31,280,664 | — | 370,318.86 / 366,753.95 |

## Tests

`tests/test_exp_ml_invoice_basis.py` — synthetic folders only (Mexico and Chile; with
`mp=True` a synthetic settlement_v2 and Pagos_Facturas): the invoice basis, actual vs
estimated refunds, coupons, dispute = return fee, withheld tax, accrual, the VAT proof and
per-country policy, the delivery-year rule, three dates, the cost-table hook, the
materiality walk, the cash bridge, the no-MercadoPago path, stability of the new lines,
the command lines, every "should be 0" row after the recalculation.
