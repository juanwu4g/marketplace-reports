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
  store_config.py      reads config.json: allowlist, batches, regions
  ziniao_client.py     Ziniao client lifecycle, HTTP IPC, open/close store, Selenium attach
  site_time.py         "today" in the selling country's timezone
  runlog.py            the six-column run log, run records, stop requests, exit-code wording
  download_diag.py     the "why was this not downloaded" step (cause codes, evidence)
  pending_store.py     reports requested but not yet collected ("owed")
  run_monthly.py       THE runner: lock, plan, child processes, summary, alert file
  app.py            the window (tkinter glue only)
  entrance.py, region_choice.py, win_tasks.py, log_view.py, ui_style.py
                       window logic that is testable without tkinter

platforms/
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
    report_build.py    the workbook itself (10 sheets, checks, bridges)
    transform.py       cleaning step (Mexico only; calls downloads/mx_sales)
    list_stores.py, open_store.py, download_history.py, pl_billing_basis.py, run_report.py

downloads/             raw exports per store and run (git-ignored, contains PII)
  mx_sales/            the Mexican cleaning package (parquet/duckdb); has its own tests
logs/                  run logs, owed list, diagnostics (git-ignored)
tests/                 the main test suite (pytest), incl. golden-file tests
docs/                  this file, the SOP, the form template, the report family
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
| **region / site** | the country of a store. In `config.json` it is `stores.regions`; inside MercadoLibre it selects the site (`mx`, `cl`, `pe`) in `meli_site.py` |
| **form** | one downloadable report (e.g. MercadoLibre "Reporte de Facturación"). Catalogued per platform in its `FORMS.md` |
| **route** | the browser steps that obtain one family of forms (billing, ventas, stock, MercadoPago) |
| **owed** | a form that was requested but not ready yet; collected later by the catch-up pass |
| **cause code** | why a form is missing (`UI_NOT_FOUND`, `PAGE_BLOCKED`, `NO_DATA` …) |
| **entrance** | who started a run: `manual` (window), `scheduled` (task), `cli` |

---

## 5. Extension points — what changes when platform #2 arrives

The move into `core/` and `platforms/` is done. These couplings are known and are
**meant to be removed while building the second platform**, not designed up front
(building one real example first is cheaper than guessing the abstraction):

| Where | Today | With several platforms |
|---|---|---|
| `core/run_monthly.py` `PLATFORM_PKG` | always `platforms.mercadolibre` | pick the package per store's platform; run platforms one after another like countries today |
| `core/store_config.py` regions | a store has a country | a store has (platform, country); Ziniao's platform name can fill the platform |
| `core/download_diag.py` | knows MercadoLibre form ids (`STOCK_FORMS`, MercadoPago kinds, `NOT_IN_REPORT`) | each platform declares its expected forms and which forms the report does not read; the cause codes and evidence logic stay in core |
| `platforms/mercadolibre/report_build.py` | builds the whole workbook, incl. styles and sheet layout | split: a shared **report kit** in core (styles, sheet scaffolding, check list, ⑩ to-do sheet) + per-platform rules. See `REPORT_FAMILY.md` |
| canonical tables | not yet | each platform translates its raw files into common order / fee / payout tables; one report builder draws the family layout from those |
| window | the country list is MercadoLibre's countries | country choice becomes (platform, country) |

Order of work for platform #2: follow `NEW_PLATFORM_SOP.md`; when its downloads
work, extract the report kit from `report_build.py` and build the new platform's
report on it, with MercadoLibre's golden tests proving nothing moved.

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
