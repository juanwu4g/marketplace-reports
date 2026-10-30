# The report family — what makes reports "sisters"

**Audience:** developers building a platform's report, and finance reviewing it.

Every platform's monthly workbook should be recognisable as the same family as
the MercadoLibre one: same reading order, same look, same way of saying "this
ties out / this does not / this is still missing". Content inside the sheets is
platform-specific; the frame is not.

The frame (styles, sheet names, check records, file names) is
`core/report_kit.py`; the sheet writers built on it (readme, P&L, bridges, long
tables, gaps) are `core/report_family.py` — see §6. Falabella's and Amazon's
reports are built with them. MercadoLibre's (the original reference,
`platforms/mercadolibre/report_build.py` + `platforms/mercadolibre/report/`) uses
the frame and the family's smaller writers — sheet heads, table headers, ①'s
rows, ②'s formulas as data, long-table rows from `Col` specs — and keeps its own
code where its sheets are richer than the family's whole-sheet writers; §8 says
which, sheet by sheet.

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
| `AZ` | Amazon | 亚马逊 | MX |
| `LV` | Liverpool | 利物浦 | MX |

Country codes are ISO 3166-1 alpha-2 (`MX`, `CL`, `PE`, `CO` …). Codes appear in
file names that finance keeps for years and in `config.json`: **once used, a code
never changes.** A new platform picks a new two-letter code (proposals: `TT` TikTok
Shop, `TM` Temu, `WM` Walmart, `SH` SHEIN) and registers it in
`core/markets.py`.

**Files.**

| What | Name | Example |
|---|---|---|
| store workbook | `<store>_<YYYY-MM>.xlsx` | `DAB_CL-Fala (Coco)_2026-09.xlsx` |
| combined per market | `<COUNTRY>_<PLATFORM>_全店汇总_<YYYY-MM>.xlsx` | `CL_FA_全店汇总_2026-09.xlsx`, `MX_ML_全店汇总_2026-09.xlsx` |
| alert, only when a person must act | `！报表异常_<COUNTRY>_<PLATFORM>_<YYYYMM>.txt` | `！报表异常_CL_FA_202609.txt` |
| download folder | `downloads/<store>/<YYYYMMDD_HHMMSS>[_m<YYYYMM>]/` | `…/20261016_175632_m202609/` |

MercadoLibre has two bases, written side by side every month: the bill basis
(`platforms/mercadolibre/bill_report`, the official one) adds the suffix
`_账单基准` to both names — `<store>_<YYYY-MM>_账单基准.xlsx`,
`<COUNTRY>_ML_全店汇总_<YYYY-MM>_账单基准.xlsx` — and the order basis keeps the
names above. Another platform that ever needs a second basis follows the same
rule: the main names stay, the second basis gets a suffix.

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

### Grade banner (sheet ①)

Every monthly **store** workbook carries its grade (docs/RESULTS_AND_PERIODS.md §3,
the same `level` / `reasons` as its results file) on sheet ①, so finance sees it
without opening anything else. `core/report_kit.stamp_grade(path, info)` writes
it; the combined workbook is not graded and has none.

- **Position**: row 1 of the first empty column after ①'s content, with one spacer
  column — ①'s content lives in A–C on every platform, so the banner is E1:F8
  (label column E, value column F). No row or column is inserted and no existing
  cell is touched; stamping again overwrites the same block (found by the 完整度
  cell in row 1).
- **Fields** (rows 1–8): 完整度 | grade · 原因 | reasons joined with "；" or "—" ·
  报表版本 | version id · 数据下载于 | the run folder's time (YYYY-MM-DD HH:MM, "—"
  when the folder name has none) · 报表生成于 | build time · 口径 | 账单基准 /
  订单基准 (MercadoLibre only; the row is skipped for one-basis platforms) ·
  是否最终版本 | 是（可以定稿的等级，且是最新的版本）/ 否（等级不够，只是未定稿）·
  last row, label column, small grey: where the grade levels are explained (本月总览
  「说明」) and that a rebuild makes a new version.
- **Colours**: the grade cell is 12 pt bold on `E2EFDA` (完整, 完整·有提醒), `FFF2CC`
  (暂定) or `FBE5E5` with red `C00000` text (有缺口, 失败); labels bold on `D9E2F3`,
  boxed. The sheet tab is green `375623` / amber `BF8F00` / red `C00000` by grade,
  and the file's properties carry it too (subject "完整度：<grade>", keywords =
  version, description = reasons) so Explorer / Finder show it without opening.
- **Rule**: the banner is stamped **by the runner** (report_build `--results-version`,
  `run_reports.bill_reports` with a stamp, `platform_reports.ReportRunner.safe_build`
  with a version) **after the builder saved the workbook and before LibreOffice
  recalculation** — openpyxl drops the cached formula values when it reopens and
  saves a workbook. The grade is computed with `formula_errors=False`; when the
  scan after recalculation finds formula errors, the banner is stamped again with
  the final grade (有缺口) and the workbook recalculated once more. **Builders never
  write it**, so the golden tests see no banner, and a banner that cannot be written
  only prints `[提示] 完整度标记没写上：…` — it never costs a report.

## 3. Look

| Element | Value (from `core/report_kit.py`) |
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
these categories (MercadoLibre's is `FEE_CATALOG` in `platforms/mercadolibre/report/catalog.py`); an
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

A sheet whose layout is its own can still use the pieces those writers are made
of, so it keeps the family's look without copying it:

| Piece | Where | What it writes |
|---|---|---|
| `title_rows(ws, title, sub, sub2)` | `report_kit` (`Sheet.title` on any sheet) | A1 navy title, row 24 high; grey notes in A2 / A3 |
| `header_cells(ws, row, cols, first_col, color, wrap_text)` | `report_kit` (inside `Sheet.header`) | header cells only: white bold on navy or band, centred, boxed |
| `header_row(ws, row, cols, widths)` | `report_family` | `Sheet.header` on a sheet without store columns (height 30, widths) |
| `set_widths(ws, widths)` | `report_family` | column widths from A |
| `readme_rows(ws, r, rows)` | `report_family` (inside `readme_sheet`) | ①'s body: `("S", section, "")` heads and `(tag, label, text)` rows |
| `pl_formula(formula, col, rows_at)`, `Span` | `report_family` (inside `pl_sheet`) | `=C5+C6-C9` / `=SUM(C5:C9)` from a formula written as data |
| `Col(…, align=)` + `Family.write_row(…, fill_color=)` | `report_family` | one table row from column specs; `align` (e.g. `wrap()`) per column, `fill_color` for the whole row |

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

`platforms/falabella/report_build.py` is the complete example (Amazon's
`platforms/amazon/report_build.py` is a second one, with a per-column transaction
catalogue `tx_catalog.py`, IVA-inclusive fees split into net + input IVA, and
withholdings as memo lines); the writers'
unit tests are `tests/test_report_family.py`. `tools/new_platform.py` generates a
`report_build.py` already laid out this way.

## 7. Cost of goods — the shared SKU cost table (`core/sku_costs.py`)

No platform's files contain the cost of goods, so without help every workbook stops
at "经营贡献毛利（未扣商品成本）". Finance keeps **one** table of unit costs for all
platforms; every report that knows about it fills ② 商品采购成本 and ⑦'s cost
columns from it.

**Where.** `config.json`:

```json
"reports": {"sku_costs": "sku_costs.xlsx"}
```

A relative path is relative to the repository root. The file is business data and
is never committed (`sku_costs.xlsx` is in `.gitignore`; keep it elsewhere if you
prefer and give an absolute path). **No key → no table → the workbooks are exactly
as before** (this is how the golden tests run).

**The file** — `.xlsx` (first sheet) or `.csv` (UTF-8, BOM ok), one header row;
English headers `platform, country, sku, unit_cost, currency, effective_date, note`
are accepted too; other columns are ignored:

| 平台 | 国家 | SKU | 单位成本 | 币种 | 生效日期 | 备注 |
|---|---|---|---|---|---|---|
| FA | CL | HD26782CL | 8900 | CLP | | |
| FA | CL | HD26782CL | 9400 | CLP | 2026-10-01 | 10 月起涨价 |
| AZ | MX | DAS-EK106PRO-BB | 365.50 | MXN | | |
| | MX | TB00204 | 389.24 | MXN | | 墨西哥所有平台 |

* 平台 / 国家 blank = any. A report's SKU is matched (trimmed, case-insensitive) on
  (platform, country), then (platform, any), then (any, country), then (any, any).
* 生效日期 optional; for a month the latest row effective on or before the month's
  last day is used. A row with an empty 单位成本 is "not entered yet" (template rows).
* 币种 must be the report's currency — there is **no FX conversion**.
* `python -m tools.sku_cost_template --out sku_costs.xlsx [--month YYYY-MM]
  [--market FA-CL --market ML-MX …]` writes (or merges into) the table: every SKU of the month's
  newest download folders of the configured stores, with reference columns (name,
  stores, units, sales). It reads files only and never deletes a row.

**Policy.** Cost of goods = unit cost × cost units. Cost units = the units of the
same lines and month as the platform's revenue (posting / transaction basis) minus
the units returned in the month where the platform's data identifies them (they go
back to stock); each platform's README states its exact rule:

| Platform | Revenue basis | Cost units |
|---|---|---|
| MercadoLibre (MX, CL, PE) | order date (Ventas) | ⑦ 表五 销售件数: units of the month's orders − orders cancelled before shipping − returns ML's review put back on sale; returns not resellable / not reviewed yet, claims, mediations, exchanges keep their cost |
| Falabella (CL) | transaction date | settled units − returned units of the month |
| Amazon (MX) | posting date | order-line units − refunded units of the month |

Unit costs are net of recoverable VAT (IVA / IGV) whatever the platform's revenue
basis — MercadoLibre's revenue lines include VAT, the cost does not change. Unit cost = landed
cost per unit as finance enters it (purchase + freight + duties). SKUs without a
cost carry no cost; ② says the coverage. ⑨ stays in quantities — valuing stock at
cost is a later step.

**In the workbook** (only when configured):

* ② 商品采购成本 = the covered SKUs' cost, still a yellow input cell (a person may
  overwrite it); its note: file, SKUs / units covered and "其中 X% 销售额已有成本；
  其余请补录或在此直接改总额". 税前净利润 follows from the existing formula.
* ⑦ 单位成本 from the table; 商品成本 = −单位成本 ×（件数 − 退回件数）; a 成本来源
  column (file row used, or "未录入成本" — yellow cell — for the rest).
* ① the file in 三、数据来源, the policy in 七、会计处理, cost and profit in 四、本期要点.
* ⑩ checks — error: file configured but missing/unreadable, rows not usable
  (non-numeric or negative cost, bad currency/platform/country code, bad date),
  duplicate (platform, country, SKU, date), currency ≠ report currency; warn:
  coverage below 100 % (uncovered SKUs by sales); the coverage line itself when 100 %.

**For a platform's report_build** (MercadoLibre, Falabella and Amazon do exactly this):

```python
from core import sku_costs as SC

table = SC.load_configured()                       # None when not configured
items = [SC.Item(sku, name, units, returned, sales) for … in ⑦]   # one per SKU
res = SC.apply(table, PLATFORM, COUNTRY, CUR, month, items)
checks += SC.checks(table, [(store, res)], PLATFORM, COUNTRY)
res.cogs, res.line(sku).unit_cost / .cost / .status, SC.status_text(line)
SC.pl_note(res), SC.source_line(table), SC.policy_line(units_basis), SC.point_line([res], profit)
```

and `sku_list(store, folder, month) -> (platform, country, currency, [Item])` so
`tools/sku_cost_template` can list its SKUs (with `run_reports.latest_folder(store,
month, base)` — the folder the monthly report would use).

MercadoLibre's ⑦ differs from the family's one-table ⑦: the unit cost goes into
表五's M column (表一 and 表四 already pick it up by store + SKU, and ⑨ values
stock with it), 商品成本 stays `=E×M` on 表五's 销售件数 (so `Item.returned` is 0:
the excluded units are already out of E), and the source goes in an extra column AE
成本来源. ① gets the file in 一、覆盖范围, the cost and profit at the end of 六 and
the policy as a new section 七、商品成本（会计处理）(重要提示 becomes 八); ⑩'s to-do
"补齐商品采购成本" becomes "补录商品成本表" (only while SKUs are uncovered).

## 8. MercadoLibre on the family

MercadoLibre's builder predates `report_family`; the family's whole-sheet
writers were distilled from it and simplified, so its sheets are richer than
what `readme_sheet` / `pl_sheet` / `bridge_sheet` / `gaps_sheet` write. It was
moved onto the family as far as the workbook could stay **identical** — every
cell, style, width, merge and the XML itself, checked on the golden store-months
of the three countries, multi-store workbooks and synthetic folders — and no
further: changing what finance already reads is a separate decision.

Code: `platforms/mercadolibre/report_build.py` (command line, assembly, recalc,
metrics; every old `report_build.<name>` still works) on top of
`platforms/mercadolibre/report/`: `site` (country profile; other modules read
`site.NAME`, so values follow `set_site`) → `catalog` → `loading` → `derived`
(`derive()`, the accounting) → `checks` → `sheet_*` (one module per sheet, ⑤⑥
together). `tests/test_ml_report_layout.py` holds the layering.

| Sheet | On the family | Its own, and why |
|---|---|---|
| ① | `readme_rows` for the body, `set_widths` | head (title 15 pt in A1, row 26, Ventas-incomplete banner merged A3:C3) and the red "注 n" notes — `readme_sheet` puts a 16 pt title in B1 and a single warning line |
| ② | subtotal / computed formulas as data (`PL_FORMULAS`, `Span`, `pl_formula`); `Sheet` | the grid: 合计 column even for one store, an extra 占GMV% column, header in row 4, row-kind colours and fonts, the memo block with its sensitivity row — `pl_sheet` drops 合计 for one store and has neither |
| ③ | `Sheet.title`, `header_cells`, `set_widths` | the bridges: numbered lines, values given as formulas, check rows in green, navy merged section bands, the MercadoPago per-order block — `bridge_sheet` lines are plain values with four styles |
| ④ | `title_rows`, `header_row`, `header_cells`; fee table = `Col` spec + `write_row` | settlement-path summary and the ② reconciliation (merged notes, deliberately empty columns) |
| ⑤ | `title_rows`, `header_row`, `header_cells`; status table = `Col` + `write_row` | subtotal rows and the KPI block (store columns with 合计 and ratio rows) |
| ⑥ | `title_rows`, `header_cells`, `set_widths`; the three per-store tables = `Col` + `write_row` | subtotal row; the "no Full in this country" text (two lines at A2/A4, not `not_applicable_sheet`'s A3) |
| ⑦ | `title_rows`, `header_cells`, `set_widths` | five tables on one 30-column grid (31 with a SKU cost table) with cross-table formulas and the cost input column — not a single `Col` table |
| ⑧ | `title_rows`, `header_row`; trend table = `trend_cols()` + `write_row` | the rule-generated notes under the table |
| ⑨ | `title_rows`, `header_row`, `header_cells`; 表一 and 表三 = `Col` + `write_row` | 表二 (per SKU): it sets the default font only on cells without a coloured one and colours single cells by verdict |
| ⑩ | `title_rows`, `header_cells`, `set_widths` | check list with merged detail cells and its own colours, the prioritised to-do list, the credit-note blocks — a different layout from `gaps_sheet` (which also lists `run_meta.json` missing files; MercadoLibre's ⑩ does not) |

When MercadoLibre's sheets are next redesigned (a decision for finance, since
the workbook changes), the remaining bespoke parts are the place to move onto
the whole-sheet writers.

