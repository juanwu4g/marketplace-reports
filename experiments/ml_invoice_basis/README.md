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
    [--folder DIR] [--base DIR] [--out-dir DIR] [--no-recalc] [--no-stability]

# several downloads of the same month: how much each basis moved, and why
.venv/bin/python -m experiments.ml_invoice_basis.stability --store <name> --month 2026-09 \
    [--folders DIR DIR ...] [--base DIR] [--out-dir DIR] [--no-recalc]
```

* Default folder: the newest `downloads/<store>/*_m<YYYYMM>` that holds the month's
  Facturación. Default stability folders: every `*_m<YYYYMM>` of the store, oldest first.
* Output: `downloads/data/reports/experiments/` (git-ignored) —
  `<store>_<YYYY-MM>_账单口径实验.xlsx` and `<store>_<YYYY-MM>_口径稳定性.xlsx`.
* The site (MX / CL / PE) comes from each folder's `run_meta.json` (none = MX); the
  previous site is restored afterwards. `config.json` is never read. Nothing opens a
  browser or a store.
* Exit codes: 0 written; 1 written but a formula error was found after the
  LibreOffice recalculation; 2 a problem the person must fix (Chinese message).

## Workbooks

`<store>_<YYYY-MM>_账单口径实验.xlsx`

| Sheet | Content |
|---|---|
| ① 说明 | what this is, **the periods**, the files and whether each invoice is closed, what the invoice can / cannot give, methods, key numbers |
| ② 账单口径损益 | invoice-basis P&L, two columns: invoice period and calendar cut; residual check (Σ lines − Σ invoice rows = 0); "发票口径无法得到" block with the order-basis values as reference only |
| ③ 与现行订单口径对比 | line by line: current / invoice period / calendar cut / differences / reason category / explanation; contribution check; order ↔ billed-sale reconciliation and an exact GMV-difference split |
| ④ 费用明细 | every FEE_CATALOG line: charged, reversed, net; unregistered Detalle values in red rows |
| ⑤ 销售与冲销明细 | reversals with link method and ratio; sales billed but not orders of the month; orders of the month not billed in the period; amount mismatches; every billed sale |
| ⑥ 口径稳定性（多次下载） | the stability table when sibling downloads of the month exist |

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

## Tests

`tests/test_exp_ml_invoice_basis.py` — synthetic folders only (Mexico and Chile),
the stability diff, the command lines, the recalculated workbook's residuals.
