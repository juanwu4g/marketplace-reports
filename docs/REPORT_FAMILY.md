# The report family — what makes reports "sisters"

**Audience:** developers building a platform's report, and finance reviewing it.

Every platform's monthly workbook should be recognisable as the same family as
the MercadoLibre one: same reading order, same look, same way of saying "this
ties out / this does not / this is still missing". Content inside the sheets is
platform-specific; the frame is not.

The reference implementation is `platforms/mercadolibre/report_build.py`. Today
the styles and sheet scaffolding live in that file; they will move into a shared
report kit in `core/` while the second platform is built (ARCHITECTURE §5).

---

## 1. Names — codes, workbooks, alerts

The naming logic lives in code in two places, and this table is its description:
`core/markets.py` (the codes) and `core/report_kit.py`
(`store_workbook_name`, `combined_workbook_name`, `alert_file_name`). New code
asks those functions; it never spells a name itself.

**Codes.** A market is `<PLATFORM>-<COUNTRY>`:

| Platform code | Platform | Chinese name | Countries today |
|---|---|---|---|
| `ML` | MercadoLibre | 美客多 | MX, CL, PE |
| `FA` | Falabella | Falabella | CL |

Country codes are ISO 3166-1 alpha-2 (`MX`, `CL`, `PE`, `CO` …). Codes appear in
file names that finance keeps for years and in `config.json`: **once used, a code
never changes.** A new platform picks a new two-letter code (proposals: `TT` TikTok
Shop, `TM` Temu, `WM` Walmart, `AZ` Amazon, `SH` SHEIN) and registers it in
`core/markets.py`.

**Files.**

| What | Name | Example |
|---|---|---|
| store workbook | `<store>_<YYYY-MM>.xlsx` | `DAB_CL-Fala (Coco)_2026-09.xlsx` |
| combined per market | `<COUNTRY>_<PLATFORM>_全店汇总_<YYYY-MM>.xlsx` | `CL_FA_全店汇总_2026-09.xlsx`, `MX_ML_全店汇总_2026-09.xlsx` |
| alert, only when a person must act | `！报表异常_<COUNTRY>_<PLATFORM>_<YYYYMM>.txt` | `！报表异常_CL_FA_202609.txt` |
| download folder | `downloads/<store>/<YYYYMMDD_HHMMSS>[_m<YYYYMM>]/` | `…/20261016_175632_m202609/` |

Historical exceptions kept for MercadoLibre (scheduled tasks and finance's folders
rely on them): the Mexican alert is `！报表异常_<YYYYMM>.txt` and the other
countries' `！报表异常_<COUNTRY>_<YYYYMM>.txt`, without the platform code.

Store workbooks and the combined one have the same sheets; the combined one has
one column per store plus a total.

## 2. Sheets — the reasoning chain

Left to right: conclusion first, then the proof, then what is still missing.

| # | Sheet | Must contain on every platform |
|---|---|---|
| ① | 阅读说明 | how to read it, which files it was built from (names + dates), what is estimated, the accountant notes for this country |
| ② | 合并损益表 | P&L from GMV down to margin before product cost; **fixed row positions** so months can be compared; original platform term next to each Chinese line; product cost as a yellow input row |
| ③ | 平账勾稽表 | the bridges: orders ↔ money received, fee invoice ↔ fees paid, tax withheld (or "nothing withheld" check). Each bridge ends in a residual that should be 0 |
| ④ | 平台费用明细 | every fee type in the period, original name + Chinese, which settlement path, reconciliation to ② |
| ⑤ | 退货退款分析 | refunds and cancellations by status, size, responsibility |
| ⑥–⑨ | platform-specific analyses | MercadoLibre: returns QC, SKU profitability, monthly trend, inventory. A platform without warehouse data says so on the sheet ("不适用：…") rather than dropping the sheet — the numbering stays stable |
| ⑩ | 数据缺口与待办 | every failed check, every missing file with its cause, every question for a person; empty means "nothing to do" |

Keep the ten-sheet numbering even when a sheet is "not applicable": people learn
where things are.

## 3. Look

| Element | Value (from `report_build.py`) |
|---|---|
| font | Arial 10; titles larger, bold |
| title band | navy `1F3864`, white bold text |
| section band | blue `4472C4` |
| sub-header | light blue `D9E2F3` |
| formula rows | grey `EDEDED` |
| totals | `FCE4D6` |
| input cells (person fills in) | yellow `FFFF00` |
| check passed / failed / warning | `E2EFDA` / `FBE5E5` / `FFF2CC` |
| text red / green / amber / grey | `C00000` / `375623` / `BF8F00` / `595959` |
| money | `#,##0.00;(#,##0.00);-` — decimals follow the currency (CLP has none) |
| integers / percent | `#,##0;(#,##0);-` / `0.0%;(0.0%);-` |
| borders | thin `BFBFBF` |

Language: sheet and line labels in Chinese; the platform's own term next to it in
the original language, so anyone can find it in the platform's files.

## 4. Checks

Every check has a key, a store, a short name, a result, a one-line detail in
plain Chinese, and a level:

- **error** — the numbers cannot be trusted until this is fixed (a bridge does
  not tie, a required file is missing, an unregistered fee type)
- **warn** — look at this, the numbers still stand (a per-order comparison with
  explained differences, a bill not closed yet)
- **skipped** — not applicable here, with the reason (e.g. no warehouse service)

The run log prints every check; the summary counts failed errors and warnings;
⑩ lists them with what to do. A report with failed errors is still written (with
a banner) — finance decides, the software does not hide data. Only `--strict`
refuses to write it.

## 5. Common data the reports are built from (target)

To keep platforms sister-like without copying code, each platform will translate
its raw files into the same three tables, and the shared report kit draws ② ③ ④
from them:

| Table | One row per | Key columns |
|---|---|---|
| orders | order line | order id, order date, status, SKU, quantity, gross amount, buyer-paid shipping, seller discounts, refunds, currency |
| fees | fee charge | fee id, charge date, order id (if any), original fee name, **category**, amount, settled how (deducted from the order / invoiced monthly) |
| payouts | money movement | date, type (sale, refund, fee, transfer to bank …), order/payment id, amount |

Fee categories (fixed list, extend with care): `commission`, `shipping`,
`ads`, `fulfilment_storage`, `penalty`, `subscription`, `tax_withheld`,
`other`. Each platform keeps a fee catalogue mapping its original fee names to
these categories (MercadoLibre's is `FEE_CATALOG` in `report_build.py`); an
unmapped fee name is an **error** check, never silently "other".

This section is the target design; MercadoLibre will be moved onto it with its
golden tests proving that no number changes.
