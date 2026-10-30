# Monthly results, grades, final versions and period reports

**Audience:** developers. The owner's rules (2026-10-22) are in §1; everything
else implements them. Folder locations come from `core/layout.py`
(`docs/DATA_LAYOUT.md`); the file format and the read side from
`core/results.py`.

## 1. The owner's rules

1. Every monthly report version gets a grade; only a complete grade counts.
2. When several complete versions exist for a month, the latest generated one is
   final. No person picks.
3. Period reports (quarterly, annual, custom) are built from the final monthly
   results — never again from platform data.
4. A new window tab, 期间报表, generates period reports.
5. The scheduled monthly run may run on several days of the month; every run
   covers all stores.
6. Once a store-month has a complete version, the other versions' files are
   deleted (new layout only).
7. A quarterly report is generated automatically after each run of the month
   that follows the quarter; the annual report after each run in January.

## 2. Results files (schema 1)

One JSON file per store × month × basis × version, written when the store's
monthly workbook is built (`core/results.write`, atomic). Path and names:
`core/results.py` docstring. Country summaries ("全店汇总") have no results
file: a period report adds up stores.

| field | type | meaning |
|---|---|---|
| `schema` | 1 | |
| `platform`, `market`, `country` | `"ML"`, `"ML-MX"`, `"MX"` | |
| `currency`, `decimals` | `"MXN"`, 2 | |
| `store`, `month` | `"BOCINA_TA02"`, `"2026-09"` | |
| `basis` | `"bill"` / `"order"` (ML), `"main"` (others) | |
| `run` | `"20261021_233002_m202609"` | the download run folder's name |
| `built_at` | `"2026-10-21T23:58:01"` | when the workbook was generated |
| `version` | `"<run>__<YYYYMMDD_HHMMSS of built>"` | one stamp per `run_reports` invocation, shared by both ML bases |
| `source_folders` | `[abs path]` | the download folder(s) read |
| `workbook` | abs path or `null` | this version's store workbook |
| `code` | `{"commit", "fingerprint"}` | `core/runlog.code_version` |
| `grade` | `{"level", "reasons": [zh]}` | §3 |
| `checks` | `[report_kit.check]` | this store's checks |
| `missing` | `[{"form", "cause", "in_report"}]` | forms not downloaded (run_meta.json) |
| `flags` | `{"provisional": [bill names], "data_incomplete": bool}` | |
| `pl` | `[{"key","label","original","kind","value"}]` | every ② row in order; `value` is the store's ② figure (computed rows evaluated), ② sign; input rows ("in") the value entered by the cost table, else 0 |
| `memo` | `[{"key","label","value"}]` | ②'s 备查项 that are amounts (no ratios) |
| `fees` | `[{"name","label","pl_key","amount"}]` | ④ by fee type, ② sign |
| `returns` | `[{"key","label","count","amount"}]` | ⑤ by class, ② sign |
| `sku` | `[{"sku","title","units","units_returned","net_sales","fees","contribution","cost"}]` | ⑦, money in ② sign, `cost` null when unknown |
| `inventory` | `null` or `{"as_of", "rows": [{"sku","title","stock","sold_30d"}]}` | ⑨ snapshot |

pl keys: data rows keep their row key; computed rows have keys too (all bases:
`contrib`, `profit`). `sku.cost` is positive (as ②'s 商品采购成本). Falabella
inventory adds `"sold_window_days": 28`. Writers: `platforms/mercadolibre/
results_files.py` (both ML bases), `core/platform_reports.py` (Falabella,
Amazon); `tools/backfill_results.py` builds them for months reported before them
(dry-run by default, `--apply`; the legacy layout writes those workbooks to
`<reports>/_backfill/<month>/` so the reports sent at the time stay as they are).

Rules for writers: the figures are the workbook's own (same objects that
filled the sheets), so a results file and its workbook never disagree; a writer
that cannot fill an optional field leaves it as an empty list / null, never
guesses. `validate()` refuses a malformed document.

## 3. Grades

Computed per store-month-basis-version from the build result, the checks and the
download folder's run_meta.json (`core/grading.py`):

| level | when |
|---|---|
| `failed` | no workbook was written |
| `gaps` | a form the report reads was not downloaded (cause other than NO_DATA); or a **blocking** check failed; or the platform flagged a source export incomplete (ML Ventas "No pudimos mostrar algunos datos", order basis) |
| `provisional` | none of the above, but a bill the month uses has not closed; or the download's completeness is unknown (`run_meta.json` missing or without `expected`: a store run cut off by the deadline, a stop or a crash — `core/grading.download_unknown`) |
| `complete_warn` | none of the above, but a non-blocking check failed |
| `complete` | nothing failed |

Blocking = error-level checks, except the ones listed in
`core/grading.NON_BLOCKING` with the reason each does not make the money figures
wrong (e.g. SKU-mapping coverage: the totals are right, only ⑦'s attribution is
less precise). That list is policy: every entry is documented there.

Missing forms: a record counts when the report reads the form (run_meta
"in_report", or `core/download_diag.affects_report` when absent) and its cause is
not NO_DATA or NOT_AVAILABLE (nothing to download). A missing form listed in
`core/grading.NON_BLOCKING_FORMS` makes `complete_warn`, not `gaps`: stock forms
(every platform; ⑥ ⑦ ⑨ only) and MercadoPago forms for the ML bill basis (status
evidence, SKU mapping and memo only). Without that list some stores could never
have a final version (EWTTO_SM has no MercadoPago report permission). Formula
errors after recalculation are gaps. `complete_warn` also comes from error-level
checks listed in `NON_BLOCKING` (keyed by platform, basis, check); warn-level
checks never change the grade. A build that wrote no workbook leaves a results
file with grade `failed` and no figures.

`reasons` are short Chinese sentences, shown in 版本说明.txt and the period
report's ①/⑩.

## 4. Final versions, the month folder and clean-up (`core/finalize.py`)

Run at the end of every report step (each platform's `run_reports`), for the
stores and month of that step.

* final(store, month, basis) = newest countable version (`core/results.final`).
* **New layout** (`layout.mode() == V2`):
  * store workbooks are written to `<monthly_dir>/_versions/<version>/`;
  * the final version's workbook is copied to `<monthly_dir>/` under the usual
    name; a store-basis without a final gets its newest version copied as
    `<name>_未定稿.xlsx` (removed once a final exists);
  * the country summary is built from each store's final version (newest version
    where there is none); if any store is not final, or a configured store has
    no version at all, its name gets `_未定稿` (the reason in 版本说明.txt);
  * `<monthly_dir>/版本说明.txt` lists, per store and basis: the final version,
    its grade, and why each other version is not final; and what was cleaned;
  * clean-up, for a store-month that has a final version for **every basis
    that has versions** (not only the bases this step built): delete the other
    versions' workbooks in `_versions/<version>/` for the bases this step built
    (the folder is shared by every store of that run and goes once empty),
    their results files, and the store-month's raw run folders that no final
    uses; drop `logs/pending.json` entries for deleted folders. A raw run
    folder is **kept** when a final uses it (by path, or by run name — results
    files written before a migration kept old paths; `migrate_layout` now
    rewrites them), when any version of the store's other months reads it,
    when it is newer (by the time in its run name) than every final and has no
    countable version (a re-download whose report failed or has gaps must
    survive until the report is fixed), or when any platform's `pending*.json`
    still owes a report in it (an unreadable pending list keeps every folder).
    Months without a final are never cleaned, nor is a store whose final could
    not be copied into the month folder (`publish_failed`: listed in
    版本说明.txt, the report step exits 1).
  * a run of a few stores still rebuilds the whole 全店汇总 from every store's
    final, so `run_monthly` does not pass `--no-combined` for store subsets in
    the new layout (it still does in the old one).
* **Legacy layout:** results files and grades are written; workbooks stay flat as
  before; no copies, no `_versions`, no clean-up.

## 5. Period reports

`run_period.py` (root) and the window's 期间报表 tab; builder in
`core/period_report/`.

* Period: `--quarter 2026-Q3`, `--year 2026`, `--from 2026-07 --to 2026-11`.
* Scope: `--region` with the units of run_monthly (`MX`, `CL`, `PE`, `FA-CL`,
  `AZ-MX`, `ALL`) and optional `--stores`. One workbook per market (currencies
  differ), every store of the market that has results in the period.
* Basis: `--basis` (ML: `bill` default = official, or `order`); `main` elsewhere.
* Input: `core/results.final()` per store-month only. A store-month without a
  final version is left out of every total and listed as 无完整版本 on ① (red
  banner) and ⑩; an alert file is written for the automatic runs.
* Before the store started: a store's **first data month** is the earliest
  month with any results file for it (any basis, any grade;
  `results.first_month_at`). A store-month before it is 未上线 (not started:
  the system had no data for the store yet), **not** missing — it is left out
  of the totals like any non-final month, but it does not make the report
  INCOMPLETE, raises no alert, and is shown grey ("未上线（系统从 YYYY-MM 起有数据）"
  on ①, ②, ⑧, ⑩; one ○ line per store in the summary). Only a month on or after
  the first data month without a final is a gap (✗, red, INCOMPLETE). A listed
  store with no results anywhere has no first data month: every month is missing.
* Output: `layout.period_dir(kind, period, market)`, name
  `<CC>_<PF>_<period>[_账单基准|_订单基准].xlsx`
  (e.g. `MX_ML_2026-Q3_账单基准.xlsx`, `CL_FA_2026.xlsx`).
* Sheets (report family, ten): ① period, months, the version and grade used for
  every store-month, what is missing; ② P&L, one column per month + period total
  (store sheet blocks + market total); ④ fees by type; ⑤ returns by month; ⑦ SKU
  totals over the period (by store and SKU); ⑧ monthly trend; ⑨ the last month's
  inventory snapshot (stock is not added up); ⑩ missing store-months, warnings
  carried from the monthly versions, versions changed since the previous
  generation of the same report. ③ ⑥ say what they cross-check or that the
  monthly reports hold the detail.
* Every total equals the sum of the final monthly results to the cent (tests).
* Automatic: after the scheduled/manual monthly run, when the month just
  reported is the last of a quarter, that quarter is (re)generated for the units
  of the run; when it is December, the year too. MercadoLibre automatically gets
  the bill basis (official); the order basis is generated on demand. A market
  with no results at all for the period is only logged; an incomplete one writes
  `！期间报表异常_<CC>_<PF>_<period>[_基准].txt` in the month's alert folder,
  which a later complete generation removes. The monthly run's exit code never
  changes because of a period report.
* The workbook records each store-month's version in a hidden sheet `_版本记录`,
  which the next generation reads to list changed versions.
* Changed versions: the period workbook records each store-month's version; a
  regeneration lists those that changed since the previous file of that name.

## 6. Scheduling several run days

`gui_settings.json` keeps a list of days per task (old single-day settings are
read as a one-day list); `core/win_tasks.task_xml(day=[3, 12, 26])` writes one
CalendarTrigger with several `<Day>`s. Every run covers all stores.
