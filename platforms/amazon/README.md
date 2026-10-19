# Amazon (墨西哥)

**Audience:** developers working on this platform, and finance reviewing the
report's rules. Operators use `README.md` at the repository root.

Market code `AZ-MX` (core/markets.py). Stores are declared in `config.json`:

```json
"catalog": {"<store>": {"platform": "AZ", "country": "MX"}},
"batches": {"AZ-MX": ["<store>"]}
```

| Command | What it does |
|---|---|
| `python -m platforms.amazon.run_batch --month YYYY-MM --stores "<store>"` | download the month's forms for these stores |
| `python -m platforms.amazon.run_batch --market AZ-MX --month YYYY-MM` | every store of the market |
| `python -m platforms.amazon.run_batch --collect-only --month YYYY-MM --stores …` | fetch reports that were requested but not ready |
| `python -m platforms.amazon.run_downloads --store … --only …` | one store, some routes |
| `python -m platforms.amazon.run_reports --month YYYY-MM` | per-store workbooks + `MX_AZ_全店汇总_YYYY-MM.xlsx` |

Downloads land in `downloads/<store>/<YYYYMMDD_HHMMSS>_m<YYYYMM>/` with
`run_meta.json` and `REPORT.md`; reports in `downloads/data/reports/financial/`.
Where each form is and how it is fetched: `FORMS.md`.

Status: the report (`report_build.py`, `tx_catalog.py`, `run_reports.py`) is
implemented and was checked against SM-Amazon's real September 2026 files (every
bridge ties, the summary PDF's four totals match to the cent); tests on synthetic
files in `tests/test_az_report.py`. Downloads: see `run_downloads.py`.

## What the report reads

One download folder per store and month; a file missing there is taken from a
later (or earlier) run of the same month (`_m<YYYYMM>` siblings), like Falabella.

| Form | File | Used for |
|---|---|---|
| order | `<YYYY><Mon>MonthlyTransaction.csv` | **every number** of ② and of the bridges |
| summary | `<YYYY><Mon>MonthlySummary.pdf` | ③ bridge D — Amazon's own control totals |
| invoice | `fee_<uuid>.xml` + `fee_invoices.json` | ③ bridge B — fee CFDIs per settlement period (no figure depends on them) |
| taxcert | `retenciones_<uuid>.xml` | ③ bridge E and the month's ISR (② memo) |
| stock | `stock_<YYYYMMDD>.csv` | ⑨ inventory |

A month the downloader recorded as `NO_DATA` for `order` is an empty month, not
an error. The transaction report must be the **Spanish** one: the reader finds
the header line by its first column `fecha/hora` (the 9-line preamble may change);
an English header (`date/time`) is an error check telling the operator to switch
the seller centre back to Español and download again.

The summary PDF is read with `pdftotext -layout` when it is installed, else with
`pypdf` / `PyPDF2` if one of them is in the virtualenv. Without any of them bridge D
is **skipped** (with that reason) — the report's numbers do not depend on it.
(Production Windows: install `pypdf` into `.venv` to get bridge D.)

Buyer city / state / postcode columns are dropped on load.

## When to run

From the **7th of the next month**: the transaction report and the summary are
final from the first days, the withholding certificate is stamped around the 6th
(`periods.taxcert_due`). Before that the certificate check is skipped and ②
shows the month's ISR as "待取得". The fee CFDI of the period straddling month
end appears only after that period closes; bridge B shows straddling periods as
"部分期间" (information only) and compares only periods that lie entirely inside
the month.

## Platform notes (for the report)

* The monthly transaction report covers Amazon's month in **GMT-6** and is
  indexed by posting date; its rows print the time in GMT-7 (GMT-8 in winter).
  The report never refilters by the printed time — the file is the month. Times are
  converted to UTC from each row's own `GMT±n` suffix (bridge B, weekly trend).
* Every row's `total` is the sum of its 15 money columns (checked per row).
* `ventas de productos`, `créditos de envío`, `descuentos promocionales` are net of
  IVA; their IVA has its own columns. The fee columns (`tarifas de venta`,
  `tarifas fba`, and `otro` on fee rows) are IVA-inclusive. Advertising is split by
  Amazon: net in `tarifas de otra transacción`, IVA in `otro`.
* `impuesto de retenciones en la plataforma` is the IVA Amazon withholds (8 % =
  half the IVA).
* An `Ajuste` without description and without SKU is either the **ISR withheld for
  the previous month** (one large negative line around the 11th) or "Ajuste del IVA
  del producto" (a few pesos). Rule: |amount| > 100 MXN → ISR
  (`tx_catalog.ISR_MIN`); bridge D checks the split against the summary's
  "Impuesto sobre la Renta retenido".
* FBA inventory reimbursements (`Ajuste`, "Reembolso de inventario de Logística de
  Amazon (<reason>)") can be negative ("ajuste general" = a clawback). Amazon's
  summary shows the positive ones under Ingresos ("Crédito de inventario del FBA")
  and the negative ones under Gastos ("Ajustes"); bridge D does the same.
* Amazon's summary splits the referral fee of order lines into "Tarifas de venta
  FBA" and "Tarifas de otra transacción", and the fee refunds into "Reembolsos de
  tarifa de venta", "Tarifas de administración de reembolso" and "Reembolsos de
  tarifas de otras transacciones"; the transaction report has all of them in
  `tarifas de venta`. The report keeps the transaction report's single line.

## Accounting rules

Decided for a Mexican IVA taxpayer reporting under IFRS-style principles, and kept
a sister of MercadoLibre Mexico's and Falabella's reports. The workbook states the
same rules on ① (七、会计处理) and the points for the tax return (八、报税时要注意).

1. **Basis — Amazon's posting month.** Every money line comes from the monthly
   transaction report, Amazon's month (GMT-6) by posting date. Amazon posts an
   FBA order when it ships, close to delivery and the transfer of control
   (IFRS 15). No refiltering by date.
2. **Revenue net of IVA.** 商品销售收入 = `ventas de productos` of order lines;
   退货退款 = the same column of refund lines (a separate, negative row); 运费收入 =
   `créditos de envío` net of refunds; 礼品包装收入 = gift wrap; 促销折扣 =
   `descuentos promocionales` (seller-funded promotions: contra-revenue, net of
   refunds); 净销售收入 = their sum. IVA columns are never revenue.
3. **Fees net of IVA.** Amazon's fee columns are IVA-inclusive (the fee CFDIs'
   totals with IVA are what the transaction report charges): ② shows each line's
   gross ÷ 1.16 (rounded to the cent per P&L line); the rest is **input IVA
   (进项税, creditable)** — a memo line, not an expense. Advertising uses Amazon's own
   split (net + IVA), never divided again. Lines: 销售佣金 (referral fee, net of
   fees refunded; the refund administration fee Amazon keeps is inside this net),
   FBA 配送费, FBA 月度仓储费, FBA 长期仓储费, FBA 入仓运输与移除费 (inbound
   transport, partnered carrier, removal/disposal), 广告费, 订阅费, 其他平台费用
   (regulatory fee, `tarifas de otra transacción` on order lines — zero so far; the
   regulatory fee comes with its own tax column, taken as net + input IVA), and
   "未登记科目 ⚠".
4. **FBA inventory reimbursements** compensate an inventory loss, they are not
   sales: 其他收入, below the fee block (as Falabella's compensation), taken as
   reported (no IVA split), clawbacks netted. ⑥ lists them line by line with the
   reason.
5. **Product cost** is not in Amazon's data: yellow input row in ②, unit-cost
   column in ⑦ (units sold per SKU are there).
6. **Taxes are memo lines (② 备查项), never expenses** — as in MercadoLibre Mexico:
   output IVA collected (product + shipping + gift wrap + promotion tax, net of
   refunds); IVA withheld by Amazon (a prepayment creditable against the seller's
   IVA); input IVA on fees (computed); IVA adjustment; an estimate of the month's
   net IVA position (output − input − withheld + adjustment); the **ISR withheld for
   this month** from the withholding certificate (ImpuestoRet 001; "待取得" without
   it); the ISR deducted in this month's cash (the previous month's lump, shown in
   the cash bridge); and a sensitivity row "若代扣税不可抵扣时的贡献毛利" =
   contribution − withheld IVA − the month's ISR.
7. **Cash.** The sum of `total` is the net change of the seller's Amazon balance;
   ③ bridge C walks from the contribution to it: + output IVA − IVA withheld − input
   IVA − ISR lump + IVA adjustment + payouts (`Trasferir`) = Σ total, residual 0.
8. **Deferred lines** (`Estado de la transacción` other than "Lanzado") stay in the
   month (posting basis) and are a warning with count and amount.

## Bridges (③)

| Bridge | Ties | Residual |
|---|---|---|
| A | order lines + refund lines (price, shipping, gift wrap, promotion) → ② net sales | error if ≠ 0 |
| B | per fee period of `fee_invoices.json`: CFDI I − E (with IVA) ↔ the transaction report's seller fees (with IVA, without advertising, which Amazon Ads invoices separately) of rows whose UTC time falls in [start, end) | warn if > max(1 %, 5 MXN); periods straddling the month are information only; payment complements (P) listed |
| C | contribution → Σ total (see rule 7) | error if > 0.05 |
| D | Amazon's summary PDF: Ingresos / Gastos / Transferencias / Impuesto ↔ the same groups rebuilt from the transaction report; ISR line ↔ the classified ISR lump | error if > 1 MXN; skipped when the PDF is missing or unreadable |
| E | withholding certificate: IVA withheld ↔ the transaction report's (timing: the certificate is by payment date `FechaServ`); the month's ISR; operations amount vs order lines (information) | warn if > max(1 % of the certificate, 50 MXN); skipped without certificate |

## Checks (⑩)

error: transaction report present (or NO_DATA), Spanish header, all columns
present, every row kind and column registered (`tx_catalog.py`), row total = sum of
columns, bridges A and C, bridge D and the ISR split, summary PDF of the right
month. warn: output IVA ≈ 16 % of product sales on order lines, IVA withheld ≈ half
the output IVA, certificate vs transaction report and certificate period, fee
invoices per closed period, stock snapshot present, deferred lines, posting days
inside the month. skipped: summary PDF missing/unreadable, certificate not yet
issued, fee periods partially in the month, no fee invoices.

## Transaction catalogue

`tx_catalog.py` maps every (`tipo`, `descripción` pattern) to a line, per money
column, with an IVA mode (net / IVA-inclusive / all IVA). Matching ignores case
and accents. A row kind that is not in the catalogue, or money in a column the
catalogue does not expect for that kind (e.g. `otro` on an order line), is an
**error** check; its amount goes to "未登记科目 ⚠" unsplit, so every bridge still
ties. Add the new kind to the catalogue once finance has decided its line.

Golden case: add SM-Amazon 2026-09 to `tests/test_golden.py` once finance has
accepted the month (docs/NEW_PLATFORM_SOP.md §5).
