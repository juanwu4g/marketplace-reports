# The report family — what makes reports "sisters"

**Audience:** developers building a platform's report, and finance reviewing it.

Every platform's monthly workbook should be recognisable as the same family as
the MercadoLibre one: same reading order, same look, same way of saying "this
ties out / this does not / this is still missing". Content inside the sheets is
platform-specific; the frame is not.

The frame (styles, sheet names, check records, file names) is
`core/report_kit.py`; the sheet writers built on it (readme, P&L, bridges, long
tables, gaps) are `core/report_family.py` — see §6. Falabella's report is built
with them; MercadoLibre's `platforms/mercadolibre/report_build.py` (the original
reference) uses the frame but still writes its own sheets — moving it onto
`report_family` is future work, to be done with its golden tests proving zero
diff.

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

## 6. Building a platform's report with `core/report_family.py`

A platform's `report_build.py` does three things: **load** its files, **derive**
one dict of numbers per store (`S = {"store": name, "m": {...}, "meta": run_meta,
…}`), **validate** (a list of `report_kit.check` records). Everything that puts
those on sheets comes from `core/report_family.py`, so the workbook is a sister
without copying layout code:

| Sheet | Writer | The platform supplies |
|---|---|---|
| ① | `Family.readme_sheet` | sections of `(label, text)`; the error banner text |
| ② | `Family.pl_sheet` + `ratio_block` + `memo_block` | a row spec of `PLRow(key, 中文, original term, kind, note, formula)`; `value(S, key)` |
| ③ | `Family.bridge_sheet` | sections of `Line(label, value(S), note, style)`; `style="residual"` on the row that must be 0 |
| ④⑦⑧⑨ | `Family.table_sheet` + `write_row` | column specs `Col(header, width, fmt, border, fill, font, formula)` and the rows |
| any sheet that does not apply | `not_applicable_sheet` | the "不适用：…" text |
| ⑩ | `Family.gaps_sheet` | the checks; missing files come from each store's `run_meta.json`; extra note sections (e.g. tax) |

Store columns and the 合计 column are handled by `Family`: with one store there
is no 合计 column; with several it follows the store columns on every sheet
(data rows sum, formula rows repeat their formula, input rows total with
`=SUM`). Money cells use the format passed to `Family(stores, money=…)` —
`report_kit.money_format(decimals)` of the currency.

P&L formulas are data, not code: a subtotal or computed row lists the rows it
adds (`"-key"` subtracts) or a `Span(first, last)` to sum a block:

```python
from core import report_family as RF
from core import report_kit as K

PL_ROWS = [
    RF.PLRow("sales", "商品销售收入", "Product sales", "d", "按结算日"),
    RF.PLRow("refund", "减：退货退款", "Refunds", "d", ""),
    RF.PLRow("net_sales", "净销售收入", "", "f", "＝ 销售 − 退款", ("sales", "refund")),
    RF.PLRow("commission", "减：佣金", "Referral fee", "d", ""),
    RF.PLRow("ads", "减：广告", "Sponsored products", "d", ""),
    RF.PLRow("fees", "平台费用小计", "", "sum", "", RF.Span("commission", "ads")),
    RF.PLRow("contrib", "经营贡献毛利", "", "f", "", ("net_sales", "fees")),
    RF.PLRow("cogs", "减：商品采购成本", "", "in", "⚠ 平台数据不含此项"),
    RF.PLRow("profit", "税前净利润", "", "f", "", ("contrib", "-cogs")),
]


def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    fam = RF.Family(stores, money=K.money_format(2))
    fam.readme_sheet(wb, K.SHEET_NAMES[0], "月度财务报表 · %s" % ctx["period_cn"],
                     "生成于 %s" % ctx["built"],
                     [("一、这份报表回答什么", [("本月赚了多少", "…")])])
    sh, r, rows_at, note_col = fam.pl_sheet(
        wb, K.SHEET_NAMES[1], PL_ROWS, value=lambda S, key: S["m"]["pl"][key],
        title="合并损益表", subtitle="口径：…")
    fam.ratio_block(sh, r + 1, "费用率（占商品销售收入）", [
        ("佣金", lambda S: -S["m"]["pl"]["commission"], lambda S: S["m"]["pl"]["sales"])],
        note_col)
    fam.bridge_sheet(wb, K.SHEET_NAMES[2], "平账勾稽表", "每座桥最后一行是残差，应为 0。", [
        ("桥 A　订单 → 结算", [
            RF.Line("订单金额", lambda S: S["m"]["orders"]),
            RF.Line("结算金额", lambda S: S["m"]["settled"]),
            RF.Line("残差（应为 0）", lambda S: S["m"]["orders"] - S["m"]["settled"],
                    style="residual")])])
    # ④–⑨: table_sheet + write_row, or not_applicable_sheet
    fam.gaps_sheet(wb, K.SHEET_NAMES[9], "数据缺口与待办", ctx["checks"])
    wb.save(out_path)
```

`platforms/falabella/report_build.py` is the complete example; the writers'
unit tests are `tests/test_report_family.py`. `tools/new_platform.py` generates a
`report_build.py` already laid out this way.
