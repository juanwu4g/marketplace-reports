# Architecture

**Audience:** whoever adds a platform or changes how the system works, human or
Claude session. Operators running the monthly job read `README.md` instead.

The system downloads seller reports from marketplace back offices (through the
Ziniao anti-detect browser) and turns them into monthly financial workbooks, one
per store plus one combined workbook per country. Today one platform is
implemented — MercadoLibre, in Mexico, Chile and Peru. The layout below is built
so that further platforms (TikTok Shop, Temu, Falabella, Walmart, Amazon, SHEIN …)
plug in next to it instead of becoming separate projects.

---

## 1. Folder layout

```
月度报表.cmd             double-click: the one window (core/app.py)
run_monthly.cmd        the runner; Windows scheduled tasks call this file
run_monthly.py         thin wrapper -> core/run_monthly.py   (kept for old commands)
list_stores.py         thin wrapper -> platforms/mercadolibre/list_stores.py
open_store.py          thin wrapper -> platforms/mercadolibre/open_store.py
bootstrap.py           checks a machine (Python, packages, Ziniao, drivers)
config.json            credentials + which stores to run   (git-ignored, secret)
launcher/              Chinese messages shown by the .cmd files (see launcher/README.txt)

core/                  shared by every platform — no platform names in here, ideally
  paths.py             ROOT = the repository root. Ask here, never use __file__.
  console.py           UTF-8 console guard for Chinese Windows
  markets.py           the registry: platforms (ML, FA …) x countries = markets ("FA-CL")
  store_config.py      reads config.json: allowlist, which market each store is, batches per market
  report_kit.py        the report family's frame: styles, Sheet, checks, names, recalc
  report_family.py     the family's sheet writers on that frame: readme, P&L from a row
                       spec (formulas as data), ratio/memo blocks, bridges, long tables,
                       not-applicable, ⑩ gaps; one store or many (+ 合计 column)
  sku_costs.py         the optional SKU cost table shared by every report (config.json
                       reports.sku_costs): read, match, checks, cost of goods per SKU
  browser_kit.py       Selenium helpers for any seller centre: shadow-root walk,
                       find/need/wait_for, js_click/real_click, table rows, download
                       folder + wait_new_files, PageError, retried()
  ziniao_client.py     Ziniao client lifecycle, HTTP IPC, open/close store, Selenium attach
  site_time.py         "today" in the selling country's timezone
  runlog.py            the six-column run log, run records, stop requests, exit-code wording
  download_diag.py     the "why was this not downloaded" step (cause codes, evidence)
  pending_store.py     reports requested but not yet collected ("owed")
  platform_batch.py    the batch runner of every platform but MercadoLibre: stores,
                       allowlist/platform checks, exit codes, summary JSON, owed catch-up
  platform_reports.py  the report runner of every platform but MercadoLibre: folder
                       choice, per-store + combined workbooks, exit codes
  run_monthly.py       THE runner: lock, plan, child processes, summary, alert file
  app.py            the window (tkinter glue only)
  entrance.py, region_choice.py, win_tasks.py, log_view.py, ui_style.py
                       window logic that is testable without tkinter

platforms/
  falabella/           Falabella Chile: same contract as mercadolibre (FORMS.md, README.md);
                       page_kit.py = its own page quirks on top of core/browser_kit;
                       report_build.py = its content on core/report_family;
                       run_batch.py / run_reports.py = thin modules on
                       core/platform_batch / core/platform_reports (amazon/ alike)
  mercadolibre/        one platform = one package
    FORMS.md           every downloaded form: where it is, how it is obtained, its columns
    README.md          platform notes: sites, quirks per country, what is not downloaded
    meli_site.py       per-country facts (hosts, locale, currency, timezone, quirks)
    meli_forms.py      seller-centre routes: billing, ventas, stock
    mercadopago.py     MercadoPago routes
    run_downloads.py   one store: request -> download -> collect -> diagnose
    run_batch.py       many stores in one Ziniao session; owed-report catch-up
    run_reports.py     builds every store's workbook + the combined one
    financial_report.py  finds the right download folder, runs report_build
    report_build.py    the workbook's entry point: command line, assembly, recalc,
                       metrics; every old report_build.<name> still resolves
    report/            the workbook by layer (each imports only the ones above it):
      site.py          country profile (currency, tolerances, tax …); set_site()
      catalog.py       raw-file structure, FEE_CATALOG, status words, number helpers
      loading.py       read one store's download folder
      derived.py       derive(): one store -> its numbers (the accounting rules)
      checks.py        validate(): the check records
      sheet_*.py       the ten sheets (readme, pl, bridge, fees, returns = ⑤⑥,
                       sku, trend, inventory, gaps) + sheet_common
    transform.py       cleaning step (Mexico only; calls downloads/mx_sales)
    list_stores.py, open_store.py, download_history.py, pl_billing_basis.py, run_report.py

downloads/             raw exports per store and run (git-ignored, contains PII)
  mx_sales/            the Mexican cleaning package (parquet/duckdb); has its own tests
logs/                  run logs, owed list, diagnostics (git-ignored)
tests/                 the main test suite (pytest), incl. golden-file tests and the
                       platform contract (test_platform_contract.py)
tools/                 developer tools, never run by the monthly job
  probe_pages.py       discovery: visit pages of one store, save screenshot + text
  probe_session.py     discovery: keep one store open, run each script dropped in a folder
  new_platform.py      scaffold platforms/<package>/ that already keeps the contract
  sku_cost_template.py list a month's SKUs into the SKU cost table for finance
                       (reads download folders only; merges, never deletes rows)
docs/                  this file, the SOP, the form template, the report family
experiments/           stand-alone analyses, never run by the monthly job (ml_invoice_basis: ML P&L on the invoice basis + stability;
                       ml_invoice_report: a ten-sheet ML report whose baseline is the billing statement)
```

The `.cmd` files, `config.json`, `downloads/` and `logs/` stay at the root on
purpose: production machines, desktop shortcuts and registered scheduled tasks
point at them.

---

## 2. How a monthly run flows

```
Task Scheduler ─► run_monthly.cmd ─► python run_monthly.py ─► core/run_monthly.py
the window (月度报表.cmd) ─────────────────────────────────────┘   (same code, --entrance manual)

core/run_monthly.py, per country, in this order:
  1. download   python -m platforms.mercadolibre.run_batch   --month M [--stores …]
  2. report     python -m platforms.mercadolibre.run_reports --month M [--region R] --out-dir …
  then: summary JSON, alert file if something needs a person, exit code
```

**Layers talk through child processes and files, not imports.** A download batch
runs for hours and drives a browser; reports are pure pandas/openpyxl. A crash in
one must never take the other down, so the runner starts each as a separate
process and reads its exit code. Downloads land on disk as the platform's own
files; reports read those files. Reports can therefore be rebuilt any number of
times without downloading again.

Child processes are started with `-m package.module` and `cwd=ROOT`. When adding a
script that the runner or the window starts, start it the same way.

---

## 3. Invariants — do not break these

1. **Scheduled tasks call `run_monthly.cmd` at the root.** Renaming or moving it
   breaks every machine that registered a task.
2. **`.cmd` files are ASCII only.** cmd.exe mis-reads non-ASCII lines after
   `chcp 65001`; user-facing Chinese goes in `launcher/*.txt` (see the README there).
3. **Never open a store that is not on the allowlist**, and never run a batch
   without an explicit store list for a non-default country. Opening a store a
   person or another run is using breaks their session.
4. **Raw downloads are evidence.** Never edit them; reports must be reproducible
   from them.
5. **Golden files.** `tests/test_golden.py` rebuilds real store-months (Mexico,
   Chile, Peru) and compares them cell by cell with the accepted workbooks. The
   data contain customer details and live outside git, in `~/report-golden/`
   (or `$REPORT_GOLDEN_ROOT`); without them the cases skip — run with `-rs` to see.
   A change that moves a number must be deliberate and approved; then
   `python tests/test_golden.py --rebaseline <case>`.
6. **Every report that is missing has a cause code** (`core/download_diag.py`).
   "Not downloaded, no reason" is a bug.

---

## 4. Concepts

| Term | Meaning |
|---|---|
| **store** | one Ziniao browser profile = one seller account in one country |
| **platform** | the marketplace (MercadoLibre, TikTok Shop …) = one package under `platforms/` |
| **market** | one platform in one country, written `ML-MX`, `FA-CL` (`core/markets.py`). Every store belongs to exactly one, declared in `config.json` `stores.catalog`; batches are per market (`stores.batches["FA-CL"]`). Old configs (`regions`, `batch`, `batches` by country) still mean MercadoLibre |
| **region / site** | the country of a store (the country half of its market); inside MercadoLibre it selects the site (`mx`, `cl`, `pe`) in `meli_site.py` |
| **form** | one downloadable report (e.g. MercadoLibre "Reporte de Facturación"). Catalogued per platform in its `FORMS.md` |
| **route** | the browser steps that obtain one family of forms (billing, ventas, stock, MercadoPago) |
| **owed** | a form that was requested but not ready yet; collected later by the catch-up pass |
| **cause code** | why a form is missing (`UI_NOT_FOUND`, `PAGE_BLOCKED`, `NO_DATA` …) |
| **entrance** | who started a run: `manual` (window), `scheduled` (task), `cli` |

---

## 5. Extension points — state after platform #2 (Falabella)

| Where | State |
|---|---|
| stores | **done** — markets (platform x country) in `core/markets.py`; `store_config` resolves a store's market, batches per market, fails closed on unknown codes or mixed batches; each platform's batch refuses other platforms' stores and cross-checks Ziniao's platform name before opening |
| `core/download_diag.py` | **done enough** — cause codes, evidence, `diagnose_store` are shared; a platform passes its own expected form ids, owed entries may carry `form`, records may carry `in_report`. MercadoLibre's form constants are still in the module (harmless) |
| report kit | **done** — `core/report_kit.py` holds the family frame (styles, `Sheet`, `check`, sheet names, workbook/alert names, recalc). MercadoLibre imports it; golden tests 0 diff |
| report family | **done** — `core/report_family.py` writes the sheets (readme, P&L from a `PLRow` spec with formulas as data, ratio/memo blocks, bridges from `Line` sections, long tables from `Col` specs, not-applicable, ⑩ gaps) for one or many stores; Falabella's and Amazon's reports are built on it. MercadoLibre's builder is split into layers (`platforms/mercadolibre/report/`) and uses the family's pieces — sheet heads, table headers, ①'s rows, ②'s formulas as data, the long tables of ④⑤⑥⑧⑨ from `Col` specs — wherever the workbook stays identical (golden 0 diff; every cell, style, layout and XML part compared on single- and multi-store MX/CL/PE builds). Its richer layouts (② grid, ③ bridges, ⑦ SKU grid, ⑩) stay its own; `docs/REPORT_FAMILY.md` §8 lists them |
| browser kit | **done** — `core/browser_kit.py`: shadow-root walk, find/need/wait_for, clicks, table rows, download folder, `wait_new_files`, `PageError`, `retried()` (the "never retry" exceptions are a parameter). Falabella's `page_kit` re-exports it and keeps only its own pop-ups, language box and downloads table. MercadoLibre's routes keep their own helpers |
| discovery | **done** — `tools/probe_pages.py` (one look per URL) and `tools/probe_session.py` (one store kept open, scripts dropped into a folder run against it); both refuse stores off the allowlist |
| new platform | **done** — `tools/new_platform.py` scaffolds a package that keeps the contract; `tests/test_platform_contract.py` checks every registered, importable platform against what the runner and the window actually start |
| `core/run_monthly.py` | **done** — runs **units**: MercadoLibre's countries keep their codes (`MX`, `CL`, `PE`) and their exact commands; other platforms' units are market ids (`FA-CL`) and start `platforms.<pkg>.run_batch / run_reports --market <id>` with an explicit store list. `--region ALL` runs MercadoLibre's countries first (same order as before), then the other units; alert per unit (`！报表异常_CL_FA_<YYYYMM>.txt`) |
| window | **done** — the unit list is `markets.all_units()`; the dropdown shows MercadoLibre's countries as before plus other platforms that have stores in config; preview, reruns and task arguments understand `FA-CL`. Old per-country scheduled tasks are only looked for for MercadoLibre |
| cost of goods | **done for MercadoLibre, Falabella and Amazon** — one optional SKU cost table for every platform (`core/sku_costs.py`, `config.json` → `reports.sku_costs`; REPORT_FAMILY §7); `tools/sku_cost_template.py` lists a month's SKUs for finance |
| runners | **done** — every platform but MercadoLibre runs on `core/platform_batch.py` (`BatchRunner`) and `core/platform_reports.py` (`ReportRunner`); its `run_batch.py` / `run_reports.py` only declare hooks (below), so a fix to the batch or report runner is made once. MercadoLibre keeps its own runners (different history: cleaning step, carried owed reports, `--region`) |
| canonical tables | not yet — Falabella's report reads its own files directly; moving both platforms onto common order / fee / payout tables stays a later step |

A platform package provides: `run_batch` (same CLI: `--month`, `--stores`,
`--market`, `--collect-only`, `--dry-run`; exit codes 0–4 — 3 store not allowed,
4 store of another platform / bad config; last line `[BATCH_SUMMARY_JSON]`),
`run_reports` (`--month`, `--stores`, `--market`, `--out-dir`, `--no-combined`,
`--dry-run`; same exit codes), `periods.form_ids(month)` / `periods.in_report(form)`
(what a store should yield and what the report reads), `FORMS.md`, `README.md`
with the accounting rules. `tests/test_platform_contract.py` is the executable
form of this paragraph; `tools/new_platform.py` generates a package that passes
it. (MercadoLibre predates part of it: no `--market`, its run_reports does not
check store platforms, its expected forms live in `download_diag` — documented
in the contract test, not changed.)

**The runner hooks.** `python -m platforms.<pkg>.run_batch` / `run_reports` stay
the entry points; the modules are thin. A platform's `run_batch.py` builds
`platform_batch.BatchRunner(code, default_market, forms=, periods=, run_downloads=,
site=, log_collected=False, name_zh=None, name_en=None)`: its route modules, its
site facts, whether the owed catch-up prints the files it collected (Falabella
yes, Amazon no). If `run_downloads` has `require_implemented()`, a real run calls
it before Ziniao starts (exit 4 on `NotImplementedError` — the scaffold's
"还没实现" gate). Its `run_reports.py` builds
`platform_reports.ReportRunner(code, __package__, default_market, periods,
folder_markers=, no_folder_reason=)`: `folder_markers(month)` lists file names /
globs tried in order — the store's folder is the newest of the month holding the
first, else the newest holding the second … (Falabella:
`NewReportTransaction_*.xlsx`; Amazon: `<YYYY><Mon>*Transaction.csv`, then
`run_meta.json`; default `run_meta.json`); `no_folder_reason` is the 未参与 text.
Labels come from `core/markets.py` (`name_zh`, `name_en`) unless given. The
workbook comes from `<package>.report_build.build`. Paths (`logs/`,
`downloads/`, the default `--out-dir`) are read from `core.paths.ROOT` when used.
The thin modules keep the old function names (`resolve_stores`, `batch_verdict`,
`collect_owed`, `latest_folder`, `resolve`, `exit_code` …) as aliases.

---

## 6. Testing

```
.venv\Scripts\python -m pytest tests -q              # main suite (Windows)
cd downloads && ..\.venv\Scripts\python -m pytest tests -q   # cleaning package
```

Tests never open a store. Window code is tested against a fake tkinter
(`tests/fake_tk.py`). Real-browser work is verified on a test machine (the VM),
never on production — see `NEW_PLATFORM_SOP.md` §2.

Other docs: `README.md` (operators, Chinese), `GUIDE.md` (operations detail),
`docs/PLATFORM_PLAYBOOK.zh.md` (lessons that hold on any platform),
`docs/NEW_PLATFORM_SOP.md`, `docs/FORM_DISCOVERY_TEMPLATE.md`,
`docs/REPORT_FAMILY.md`, `platforms/mercadolibre/FORMS.md`.
