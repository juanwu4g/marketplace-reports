# MercadoLibre report on the billing-statement baseline (账单基准)

**Audience:** developers and finance.

Part of the monthly run since 2026-10 (owner and accountant: the bill basis is the
**official** one, the order-based report stays as the management view). Every
month `platforms/mercadolibre/run_reports.py` writes both bases for every ML
country (MX, CL, PE), bill basis first:

| | store | country summary |
|---|---|---|
| bill basis (this package) | `<store>_<YYYY-MM>_账单基准.xlsx` | `<CC>_ML_全店汇总_<YYYY-MM>_账单基准.xlsx` |
| order basis (unchanged) | `<store>_<YYYY-MM>.xlsx` | `<CC>_ML_全店汇总_<YYYY-MM>.xlsx` |

`run_reports --basis bill|order|both` (default both; `config.json` →
`reports.ml_basis: ["bill", "order"]` changes the default). The two bases never
import each other; one failing does not stop the other. `combined.py` builds the
country summary from the same per-store data (`build.prepare`) as the store
workbooks, so its store columns equal the store workbooks. A bill the month uses
that MercadoLibre still marks "podría tener cambios hasta el cierre" (Chile's
cut-day bills) makes the workbook **暂定** (banner on ①, warn check
`invoice_open`): regenerate with the window's 「只出报表」 after it closes.
Golden cases: `tests/test_golden.py` (`*.bill.baseline.json`, `cl.bill-combined…`).

Originally built as an experiment (`experiments/ml_invoice_report/`, 2026-10-21).

A MercadoLibre financial workbook built **only** from the billing statement —
`Reporte_Facturacion_MercadoLibre_<Mon><YYYY>.xlsx` and
`Reporte_Notas_Credito_MercadoLibre_<Mon><YYYY>.xlsx` — in the report family's
ten-sheet format (`core/report_family.py`, `docs/REPORT_FAMILY.md`). It is written
independently of the official order-based report and of the earlier
`experiments/ml_invoice_basis/`: it imports nothing from
`platforms/mercadolibre/report*`, `pl_billing_basis.py` or that experiment, and it
never shows their numbers.

```
python -m platforms.mercadolibre.bill_report --store <name> --month 2026-09 --folder <download folder> \
       [--site mx|cl|pe] [--out-dir DIR] [--sku-costs FILE] [--no-recalc]
```

* `--site`: from the folder's `run_meta.json` (`"site"`); older Mexican folders have
  none → pass `--site mx`.
* Output: `<store>_<YYYY-MM>_账单基准报表.xlsx`, default folder
  `downloads/data/reports/experiments/` (outside git). Recalculated with LibreOffice
  when `soffice` is installed.
* Exit code 0 ok, 1 error-level checks failed (the workbook is still written, with
  a banner on ①), 2 cannot run.

## Baseline and period

* **Ledger** = every row of every Facturación and Notas de crédito file in the
  folder, each charge once (`Número del cargo`), with its sign as MercadoLibre
  states it (charges positive, `Anulación …` negative), **with IVA**.
* **Period** = the calendar month by **charge date** (`Fecha del cargo`). Mexico and
  Peru bill by calendar month. Chile bills per store from a cut day (DAB_CL: the 8th
  to the 7th), so September = the Sep 1–7 rows of the "Sept" invoice + the
  Sep 8–30 rows of the "Oct" invoice + credit-note rows dated in September. A check
  fails (error) if the folder's invoices do not cover the whole month.
* **Sales** come from the same ledger: each `Cargo por venta` row carries the sale it
  charges (Número de venta, quantity, unit price, `Total de la venta`, buyer-paid
  shipping `Envío a cargo del cliente`). A sale is in the month's GMV when its
  commission row's charge date is in the month (a sale billed twice — re-billed
  after a reversal — counts each time).
* **Cancellations / returns**: the invoice does not state refunds, only commission
  reversals (`Anulación del cargo por venta`). Reversed share = reversal ÷ original
  commission (original found through `Cargo que bonifica` / `Cargo que anula`; if
  the original is in an invoice that was not downloaded, the listing's median
  commission rate in the folder estimates it, else a full reversal is assumed —
  each case is a warn check). Reversed sale amount = share × `Total de la venta`;
  ≥ 99 % is a full reversal (units come back), otherwise a partial refund (units
  stay). MercadoPago's actual refund is shown next to it as evidence, never as the
  amount.
* Reversals dated in the month count in the month even for an earlier sale; a
  month's sale reversed after month-end stays in the month's GMV and is listed as
  a post-period item (② memo, ⑤ 三-E).

## Detalle catalogue (`catalog.py`)

| Detalle | ② line | To SKU |
|---|---|---|
| Cargo por venta / Anulación del cargo por venta | 销售佣金（净） | the sale's SKU |
| Cargo por venta con afiliados (/ Anulación) | 联盟推广佣金 | the sale's SKU |
| Cargo por envíos de Mercado Libre / Anulación | 平台运费（净） | the sale's SKU |
| Cargo por devolución / Anulación | 退货处理费（净） | the sale's SKU |
| Cargo por campaña de publicidad de Product Ads / Anulación (…Products Ads) | 广告费 | allocated |
| Cargo por campaña de Publicidad de Seguidores (/ Anulación) | 广告费 | allocated |
| Cargo por campaña de publicidad de Display Ads (/ Anulación) — documented in FORMS.md, not seen yet | 广告费 | allocated |
| Cargo por servicio de almacenamiento Full | Full 仓储类费用 | storage detail by day |
| Cargo por stock antiguo en Full | Full 仓储类费用 | Full charges report / listing |
| Cargo por sobrepasar espacio Full, Cargo por comprar más espacio en Full | Full 仓储类费用 | allocated |
| Cargo por servicio de colecta Full, Cargo por retiro de stock Full | Full 揽收/取回/违约 | Full charges report / listing |
| Cargo por incumplimiento en Envíos Full | Full 揽收/取回/违约 | allocated |
| Cargo por mantenimiento de Mi página / Anulación | 店铺服务费 | allocated |
| Cargo debitado del dinero en garantía del Programa Beneficio de Reputación | 其他平台扣款 | allocated |

Any other Detalle is an **error** check; its amount goes to 其他平台扣款 so the
ledger still ties.

## Store-level costs (policy, stated on ①)

1. Charges tied to a sale → that sale's SKU.
2. Full collection / withdrawal / old-stock charges → SKU from
   `Reporte_Cargos_Full` (by charge number) or the row's listing id.
3. Full daily storage: the invoice charge of day *d* is the storage of day *d−1*
   (verified to the cent against `Costos_por_servicio_almacenamiento`); it is split
   by that day's per-SKU amounts in the storage report, or the nearest day's when
   the report does not cover it.
4. Everything else (ads, Mi página, Full space and penalties, reputation programme,
   unknown Detalles, Full charges with no SKU) → allocated by each SKU's net GMV
   (negative counted as 0), in separate "分摊" columns of ⑦.

## SKU mapping (`mapping.py`)

T1 sales export (`# de venta` = invoice sale number, or the pack number for a
single-item pack, or the pack head's item with the same listing) → T1b
MercadoPago collection (`order_id` / payment with `SKU Producto`) → T2 listing id
→ SKU from every file pairing them (unique; or unique by unit price among a
listing's variants) → T2* most frequent SKU of the listing → T3 title. Each row of
⑦ 表二 shows tier, confidence and source; unmapped sales are listed on ⑩.

## Status (`status.py`)

Partial refund (< 99 %); returned with refund (return/claim/mediation status, ML
review, return fee, warehouse triage, buyer-protection refund after dispatch, or any
shipping/delivery evidence); cancelled before fulfilment (cancel status, or
MercadoPago cancelled/refunded by admin/payer with no shipment; with no evidence at
all, a reversal ≤ 3 days after the sale with no return fee — marked 推定);
reversed with no evidence; completed; open/unknown (in transit, open claims, refund
evidence not yet reversed in the invoices — a warn).

## Checks

Error: invoices do not cover the month, unknown Detalle, any residual ≠ 0, SKU
mapping below 95 % of billed GMV, unreadable file. Warn: closing invoice not updated
after month-end, a bill month without credit notes, duplicate charge numbers, sale
billed twice, reversal share estimated, reversals without evidence, refunds not yet
reversed, mapping below 99 %, storage attribution by nearest day > 25 %, Full
charges without SKU, invoice not (fully) paid in `Reporte_Pagos_Facturas`, credit
notes not matched. Skipped: no MercadoPago files, no Full (Peru).

## Files

`profile.py` country facts · `catalog.py` Detalle catalogue · `readers.py` file
readers · `mapping.py` SKU tiers · `status.py` status classes · `model.py` load,
derive, residuals, checks · `workbook.py` the ten sheets · `build.py` one
store-month · `__main__.py` CLI. Tests: `tests/test_exp_ml_invoice_report.py`
(synthetic folders only).
