# meli-mx-pipeline — 运维指南 / Operations Guide

Automating the 紫鸟 (Ziniao) SuperBrowser to pull sales data from MercadoLibre MX
and other stores.

> **Status:** all four routes (billing, ventas, stock, mercadopago) verified
> end-to-end against `BOCINA_SM` on 2026-09-02. See platforms/mercadolibre/FORMS.md for the 15 forms. Sections marked **TODO** are scaffolding to be filled in
> as the extraction work continues.

---

## 1. Quick start on a new machine

```bash
git clone <repo>                 # or copy the folder
cd mx_sales_data

python -m venv .venv
.venv\Scripts\pip install -r requirements.txt      # Windows
# source .venv/bin/activate && pip install -r requirements.txt   # mac/linux

copy config.example.json config.json               # then fill in credentials
.venv\Scripts\python bootstrap.py                  # verify the machine
```

`bootstrap.py` exits non-zero and prints a fix for anything missing. Get it to
`environment OK` before running anything else.

⚠️ **On Chinese Windows, run `set PYTHONUTF8=1` before the first `pip install`.**
The pip bundled with Python 3.11.9 reads `requirements.txt` in the system
encoding (GBK), and the file carries UTF-8 Chinese comments, so a fresh install
dies with a GBK decode error before installing anything. Verified on a fresh
machine.

---

## 1.5 Which command do I run?

**The one command that does everything, start to finish, unattended:**

```bash
.venv\Scripts\python -m platforms.mercadolibre.run_batch --workbooks
```

That opens every store in `config.json` → `stores.allowed`, one at a time,
downloads all four routes from each, cleans the whole tree, and writes one
accounting workbook per store to `downloads/data/reports/<STORE>_<stamp>.xlsx`.
No human input at any point.

Drop `--workbooks` for the raw files plus cleaned tables, without the workbooks.

| Command | Downloads? | Cleans? | Workbook? |
|---|---|---|---|
| `python -m platforms.mercadolibre.run_batch --workbooks` | **yes**, every allowed store | **yes** | **yes**, one per store |
| `python -m platforms.mercadolibre.run_batch` | yes | yes | no |
| `python -m platforms.mercadolibre.run_batch --no-transform` | yes | no | no |
| `python -m platforms.mercadolibre.run_batch --dry-run` | no | no | no |
| `python -m platforms.mercadolibre.run_batch --collect-only` | **only what's owed** | yes | no |
| `python -m platforms.mercadolibre.run_downloads` | **one** store only | no | no |
| `cd downloads && python -m mx_sales build` | no | yes | no |

**These are the building blocks.** A production run — either window, the
scheduled task, or `run_monthly.cmd` by hand — goes through one runner,
`run_monthly.py`, which calls `run_batch.py --month …` and then `run_reports.py`
and owns the lock, the run log, the alert file and the download deadline. See
§10.5.

Add `--stores NEW_STORE` to any of the batch forms to run one store only —
that is the whole recipe for onboarding a new store, once its name is in
`stores.allowed`:

```bash
.venv\Scripts\python -m platforms.mercadolibre.run_batch --stores NEW_STORE --workbooks
```

### The two things people get wrong

**`stores.allowed` is a permission list, not a work list.** Putting two stores
there does not make `run_downloads.py` download two — that script is
single-store by design and always runs exactly one (`stores.default`, or
`--store`). Downloading several is what `run_batch.py` is for. It warns you:

```
note: running BOCINA_TA02 only. 1 other store(s) allowed (UNIT_TA04) - use run_batch.py
```

**`run_downloads.py` does not clean.** Only `run_batch.py` runs the transform.
After a single-store run, clean by hand with
`cd downloads && python -m mx_sales build`.

### About the workbooks

`--workbooks` runs the cleaner's `validate` once per store, on that store's
ventas export from **this** run.

**Each workbook is written the moment its own store finishes**, not batched at
the end of the run. A store that finishes at 09:10 has a usable workbook at
09:10 rather than three hours later, which is what matters when a run is
interrupted: the work already done is already delivered.

That is possible because `validate` reads the RAW export and pairs it with the
facturación files in the SAME run folder - it needs neither the Parquet build
nor any other store. The independence is what makes per-store generation
correct, not merely convenient.

The one exception is a store whose ventas missed the in-session budget and was
collected late; a catch-up pass after the late collect writes its workbook, so
it does not silently end up without one.

Three further things follow from how `validate` works:

- It reads the **raw** export and pairs it with the facturación files in the
  **same run folder**, so each workbook is consistent to one snapshot.
- It takes one file, so a store with no ventas export in that run gets no
  workbook and says so rather than failing.
- It no longer waits for the cleaning step, since it does not read the cleaning
  step's output.

A workbook that cannot be written (most often because it is open in Excel)
is recorded as a store error; the downloads and cleaned tables are unaffected.

### What a run produces

```
downloads/<STORE>/<YYYYmmdd_HHMMSS>/    the raw files, plus REPORT.md describing them
downloads/<STORE>/<stamp>/run_meta.json what was expected, what is missing and why
downloads/<STORE>/<stamp>/_diag/        evidence for each report that did not arrive
downloads/data/processed/<report>/      cleaned Parquet
downloads/data/mx_sales.duckdb          views over the Parquet
logs/batch_<stamp>.json                 per-store status, missing reports by cause
logs/runs/<run_id>.log / .json          one pair per run of the runner (§10.5)
```

Useful options, all shared by both entry points: `--only billing|ventas|stock|mercadopago`,
`--months N` (billing periods), `--skip-ip-check`, `--no-restart` (attach to an
already-running Ziniao client instead of restarting it).

Full detail: §9 for the routes, §10 for the batch and the cleaning step.

---

## 2. What must be installed

Two things live **outside** this repo and cannot be vendored — everything else
is either in `requirements.txt` or downloaded on demand.

| Requirement | Version verified | How to get it |
|---|---|---|
| **Ziniao client** | 6.27.1 (V6) | Install from the Ziniao console. Windows V6 binary is `ziniao.exe`; V5 is `starter.exe`. |
| **Python** | 3.11.9 | Any 3.8+; 3.11 is what's verified. |
| `requests` | 2.32.2 | `requirements.txt` |
| `selenium` | 4.48.0 | `requirements.txt` |
| chromedrivers | 18 versions | **Optional.** `python bootstrap.py --drivers` → `./webdriver/` (~300 MB). Rarely needed — see §5. |

### Account requirement

The Ziniao account needs **WebDriver permission enabled** in the console, and it
must be an **enterprise login** (企业登录 — company + username + password). A
personal login will not work with this API.

---

## 3. Files in this project

> **Where files live (since 2026-10):** shared modules are in `core/`, MercadoLibre
> modules in `platforms/mercadolibre/`; file names are unchanged. Scripts are started
> as modules, e.g. `python -m platforms.mercadolibre.run_batch`. The root keeps the
> `.cmd` files, `config.json`, `downloads/`, `logs/` and three forwarding scripts
> (`run_monthly.py`, `list_stores.py`, `open_store.py`). See `docs/ARCHITECTURE.md`.

| File | Purpose |
|---|---|
| `bootstrap.py` | Environment check / driver download. Run this first on any new machine. |
| `store_config.py` | **Which stores the tooling may touch.** Reads `config.json`; enforced in `find_store()`. |
| `list_stores.py` | Lists the account's stores; marks which are on the allowlist. |
| `open_store.py` | Opens a store for manual poking, no Ziniao login screen. |
| `ziniao_client.py` | Client lifecycle, HTTP IPC, store open/close, Selenium attach. |
| `meli_forms.py` | MercadoLibre download routes (billing, ventas, stock), incl. shadow-DOM helpers. |
| `mercadopago.py` | MercadoPago reports (6 types, 4 UI flavours): request / probe / fetch. |
| `run_downloads.py` | Single-store entry point, and `run_store()` — the pipeline both entry points share. |
| `run_batch.py` | Multi-store entry point. One client session, serial stores, isolated failures, summary JSON. |
| `run_monthly.py` | **The one runner.** Both windows, the scheduled task and the command line start it; it runs `run_batch.py` then `run_reports.py` as subprocesses. Lock, run log, summary, alert file, download deadline. See §10.5. |
| `app.py` | The one window (`月度报表.cmd`): run now, needs-human-action, monthly schedule, live log. Starts `run_monthly.py --entrance manual …` detached and registers the Task Scheduler task that runs `run_monthly.cmd --entrance scheduled --region ALL`. Logic lives in `entrance.py` / `region_choice.py` / `win_tasks.py`. |
| `runlog.py` | The run-log format (six columns), run records, "is a run in progress", stop requests, exit-code wording. No tkinter. |
| `log_view.py` | The log viewer both windows embed. Tk glue only; the logic is in `runlog.py`. |
| `download_diag.py` | Why a report was not downloaded: evidence capture, the closed `CAUSES` table, the list of expected forms. Imports neither selenium nor tkinter. |
| `run_report.py` | Writes `REPORT.md` into each run folder. Presentation only — never touches the browser. |
| `pending_store.py` | Reports requested but not yet collected (`logs/pending.json`); entries that age out are archived to `logs/pending_expired.json`. Survives a crash or reboot. |
| `transform.py` | The seam to `downloads/mx_sales`: runs its CLI as a subprocess after a batch. |
| `pipelines/mx_sales/` | **Separate project** (own README, pyproject, tests): cleans the raw exports to Parquet + DuckDB. Its data folders are passed in (`MX_SALES_*`, from `core/layout.mx_sales_paths`). |
| `core/layout.py` | Where every kind of data lives: the old layout (`downloads/`, `logs/` in the code folder) or the new one under `config.json` → `paths.data_root` (docs/DATA_LAYOUT.md). |
| `core/results.py`, `core/grading.py`, `core/finalize.py` | Monthly results files, grades, final versions, the month folder and clean-up (docs/RESULTS_AND_PERIODS.md). |
| `core/period_report/`, `run_period.py` | Quarterly / annual / custom period reports from the final monthly results; the window's 期间报表 tab. |
| `tools/migrate_layout.py`, `tools/backfill_results.py` | Move a machine to the new layout (verified, reversible); build results files for months reported before them. |
| logs (`logs/`, or `<data>/logs/` in the new layout) | **Gitignored.** `batch_<stamp>.json`, one per batch run; `runs/<run_id>.log` + `.json`, one pair per run of the runner; `current.json` while a run is in progress; `pending.json`, `pending_expired.json`; `report_diag/`. |
| `report_build.py` | The monthly financial report generator: 10 sheets, P&L → bridges → fees → returns → SKU → trend → inventory → checks. Entry point only; the workbook is built by `report/` (country profile, loading, `derive()`, checks, one module per sheet) on `core/report_kit` + `core/report_family`. Runs as a plain script too (puts the repo root on `sys.path`). |
| `financial_report.py` | The seam to `report_build.py`: picks each store's newest *complete* run folder and runs it as a subprocess. |
| `run_reports.py` | Entry point for the monthly reports: one workbook per store plus an all-store roll-up. |
| `pl_billing_basis.py` | Verification tool: rebuilds the P&L on a pure billing-date basis to compare against the current hybrid basis. Safe to delete. |
| `experiments/ml_invoice_basis/` | Experiment, not part of the monthly run: MercadoLibre P&L on the invoice basis (Facturación + Notas de crédito) vs the order basis, and their stability across downloads. See its README. |
| `platforms/mercadolibre/bill_report/` | The MercadoLibre bill basis (账单基准, the official one): a ten-sheet workbook per store and per country whose only amount baseline is the billing statement (Facturación + Notas de crédito, by charge date), with SKU mapping, cancellation/return status and store-level cost allocation. Written every month next to the order-based report by `run_reports` (`--basis`). See its README. |
| `downloads/` | **Gitignored.** Old layout only: raw output, one folder per run (new layout: `<data>/raw/`). |
| `config.json` | **Gitignored.** Real credentials and machine-specific paths. |
| `config.example.json` | Template with every field documented. Copy → `config.json`. |
| `requirements.txt` | Pinned runtime deps. |
| `.venv/` | **Gitignored.** Project-local Python environment. |
| `webdriver/` | **Gitignored.** Optional fallback chromedrivers. |
| `platforms/mercadolibre/FORMS.md` | Catalogue of every downloaded form: source page, how obtained, file structure. |
| `GUIDE.md` | This file. |

TODO — add rows as exporters / parsers land.

---

## 3.5 Which store gets opened

The store name lived in six hardcoded places until 2026-09-01. It now lives in
`config.json` only:

```json
"stores": {
  "default": "BOCINA_SM",
  "allowed": ["BOCINA_SM"]
}
```

`default` is used when `--store` is omitted. `allowed` is the **only** set of
stores that may be opened; an empty list means unrestricted.

One optional key sits beside them: `"no_full": ["STORE", …]` names stores that
have **no Full warehouse**. Their missing stock reports are then classified
`NOT_AVAILABLE` instead of needing attention (§10, "Why a report is missing").
`--no-full-stores STORE …` does the same for one invocation. The script never
guesses this: on the page, "no Full" and "the stock menu did not render" are
both an empty menu.

**Why this matters.** `getBrowserList` returns every store the enterprise
account can see — 93 on this account — and its records carry **no permission
field** (`browserId, browserIp, browserName, browserOauth, isDynamic,
isExpired, platform_id, platform_name, proxyType, siteId, store_username,
tags`). So the client cannot tell which stores are authorised for WebDriver, and
nothing in the IPC can be asked: the local API has only `getBrowserList`,
`startBrowser`, `stopBrowser`, `updateCore` and `exit` — verified against the
client bundle, the vendor demo and the official repo. WebDriver permission is
granted in the Ziniao console (docId 99), not exposed over the wire.

The allowlist is therefore the only local guard. It is enforced in
`ZiniaoClient.find_store()` — the single choke point every path to
`startBrowser` passes through — and it fails **before** any request reaches
Ziniao:

```
STORE NOT ALLOWED: store 'NARWAL_SM' is not in config.json -> stores.allowed
(BOCINA_SM). Add it there if you meant to open it.
```

Entry points exit 3 on that error, distinct from 2 for a Ziniao failure.
`python list_stores.py --mx` marks allowed stores with `*`.

---

## 4. How the connection works

```
  your script
      │  HTTP POST  {action, requestId, company, username, password}
      ▼
  ziniao.exe  --run_type=web_driver --ipc_type=http --port=16851
      │  returns debuggingPort per store
      ▼
  store browser (Chromium)  ←── selenium attaches via debuggerAddress
```

Every IPC call is a POST of a JSON body to `http://127.0.0.1:<socket_port>`,
with the credentials merged into each request. `statusCode: 0` means success;
`-10003` means a login/permission problem.

Sequence for one store:

1. Fully exit any running Ziniao client (`taskkill /f /t /im ziniao.exe`).
2. Launch it with `--run_type=web_driver --ipc_type=http --port=<port>`.
3. Poll the port until it answers.
4. `getBrowserList` → find the store by `browserName`, take its `browserOauth`.
5. `startBrowser` with that oauth → returns `debuggingPort`, `browserPath`,
   `launcherPage`, `ipDetectionPage`.
6. Attach Selenium to `127.0.0.1:<debuggingPort>` using the driver at
   `<browserPath>\webdriver.exe`.
7. Open `ipDetectionPage`, confirm the proxy is healthy.
8. **Navigate to `launcherPage`** (mandatory — see §6.3).
9. Do the work.
10. `stopBrowser` with the same oauth.

### Verified reference values

```
client        D:\ziniao\ziniao.exe   (v6.27.1)
IPC port      16851
stores        92 visible to this account
test store    BOCINA_SM
              browserOauth  05lhKoVWH4E2nWBhxRXXvA==
              browserId     27797302275609
              platform      MercadoLibre-墨西哥-本土
              exit IP       216.238.92.32  (static proxy)
              core          Chromium 146.1.4.66
              launcherPage  https://www.mercadolibre.com.mx/resumen
```

Note `browserOauth` values contain `+`, `/` and `==`. They are base64 — pass them
as JSON string values, never interpolate them into a URL.

---

## 5. Chromedrivers: usually not needed

`startBrowser` returns `browserPath`, and that folder ships its own matching
`webdriver.exe`:

```
C:\Users\<you>\AppData\Roaming\ziniaobrowser\env-kit\Core\chrome_64_146.1.4.66\webdriver.exe
```

Prefer it. It is guaranteed to match the store's core version. The 18-driver
fallback set only matters for older cores that don't bundle one. Do not commit
these binaries.

---

## 6. Gotchas — all four cost real debugging time

### 6.1 `ELECTRON_RUN_AS_NODE` breaks the client launch

Symptom:

```
D:\ziniao\ziniao.exe: bad option: --run_type=web_driver
```

VS Code's extension host sets `ELECTRON_RUN_AS_NODE=1`, and any process spawned
from it inherits the variable. That makes `ziniao.exe` (Electron 27) start as a
bare Node process, which normalizes `--run_type` → `--run-type`, doesn't
recognize it, and exits immediately.

Always strip it when spawning the client:

```python
env = os.environ.copy()
env.pop("ELECTRON_RUN_AS_NODE", None)
subprocess.Popen(cmd, env=env)
```

Without this the code works from a plain terminal and fails from inside an
editor — an easy day to lose.

### 6.2 `updateCore` may never return `statusCode: 0`

On this machine it returns `-10000 / "处理中"` (processing) indefinitely. The
vendor demo loops on it with `while True` and would hang forever. `startBrowser`
works fine regardless. Cap the retries and treat failure as non-fatal.

### 6.3 Never end a session on the IP-check page

`ipDetectionPage` is a `chrome-extension://` URL. If that is the only open tab,
Chromedriver refuses to attach on a later run:

```
unknown error: unable to discover open window in chrome
```

Always navigate to `launcherPage` after the IP check. To recover a browser
already stuck in that state, open a normal tab over CDP first:

```bash
curl -X PUT "http://127.0.0.1:<debuggingPort>/json/new?https://example.com"
```

### 6.4 Console encoding on Windows

Store names and API errors are Chinese. Printing them under the default `cp1252`
console raises `UnicodeEncodeError`. Set `PYTHONIOENCODING=utf-8`, or write logs
to a UTF-8 file rather than stdout.

---

## 7. IPC actions reference

| Action | Purpose | Key fields returned |
|---|---|---|
| `getBrowserList` | List all stores on the account | `browserList[]` with `browserName`, `browserOauth`, `browserId`, `browserIp`, `platform_name`, `tags`, `isExpired` |
| `startBrowser` | Open one store | `debuggingPort`, `browserPath`, `launcherPage`, `ipDetectionPage`, `downloadPath`, `core_version` |
| `stopBrowser` | Close one store | — |
| `updateCore` | Pre-download cores | see §6.2 |
| `exit` | Shut the client down | — |

Useful `startBrowser` options: `isHeadless`, `privacyMode`,
`isWebDriverReadOnlyMode`, `cookieTypeSave`, `injectJsInfo`.

TODO — document the option combinations we settle on for unattended runs.

---

## 8. Target stores

93 stores are visible on the account. **Visibility is not permission** — the
tooling only touches stores on an explicit allowlist, held in `config.json`:

```json
"stores": { "default": "BOCINA_SM", "allowed": ["BOCINA_SM", "EWTTO_SM"] }
```

`store_config.py` is the single source of truth and `ZiniaoClient.find_store()`
is the only place it is enforced, so every entry point inherits the same scope.
`python list_stores.py` prints the account's stores and marks allowed ones `*`.

The allowlist **fails closed**. An earlier version returned `{}` on a JSON parse
error, which made `allowed` fall back to `[]` — read as "unrestricted" — so one
missing comma would have silently unlocked all 93 stores. A malformed
`config.json` now raises `StoreConfigError` and exits 4.

| Store | Platform | Notes |
|---|---|---|
| `BOCINA_SM` | MercadoLibre MX | Verified end to end, all four routes |
| `EWTTO_SM` | MercadoLibre MX | On the allowlist; routes not yet run |

Ziniao exposes no API for *which* stores WebDriver is authorized on —
`getBrowserList` returns everything on the account. The allowlist is ours, not
the vendor's.

---

## 9. Data extraction

Four download routes are implemented — three from MercadoLibre's left-tab menu
and one from MercadoPago. Entry point: `run_downloads.py`. **platforms/mercadolibre/FORMS.md is the
per-form reference**; this section covers only the mechanics.

```bash
.venv\Scripts\python -m platforms.mercadolibre.run_downloads                      # all four routes
.venv\Scripts\python -m platforms.mercadolibre.run_downloads --only ventas        # route 2 only
.venv\Scripts\python -m platforms.mercadolibre.run_downloads --only billing --months 2
.venv\Scripts\python -m platforms.mercadolibre.run_downloads --skip-current       # skip the EN CURSO month
.venv\Scripts\python -m platforms.mercadolibre.run_downloads --store EWTTO_SM     # must be on the allowlist
```

### The run has three phases

Some reports generate server-side for minutes. Rather than block on each, the
run **requests** them first (phase 1), does everything that downloads instantly
(phase 2), then **collects** (phase 3).

Phase 3 is **one shared wait**, cycling over everything still outstanding until
all are collected or `--collect-timeout` (default **1800s**) expires. An earlier
version divided the budget between pending items — with seven reports and 600s
that gave each 100s, and nothing taking 2.6–6.5 minutes could finish in its
slice. Raising the timeout was not the fix; sharing it was.

Measured 2026-09-02: all 7 pending reports were ready on **cycle 1**, because
phase 2 absorbed the generation time. Anything still not ready is listed under
`[pending]` and the run continues — the report stays on its page and a later run
picks it up, so a slow generation never fails the run.

Files land in `downloads/<STORE>/<timestamp>/`. Downloads are redirected there
per-session over CDP (`Browser.setDownloadBehavior`); without that they go to
the store's own `downloadPath`, which is shared across runs.

⚠️ **The path must be absolute.** CDP resolves a relative `downloadPath` against
the *browser process's* working directory, not ours — `--out downloads/x` made
every route click correctly and then report `got 0 files`, with the files
nowhere findable. `run_downloads.py` runs `--out` through `os.path.abspath()`.

### Route 1 - Facturación reports

`Tarifas y pagos` -> `/billing/resume` (page title: **Facturación**) -> newest
N month cards -> `Ir al detalle` -> `Reportes` tab -> `Seleccionar todos los
reportes` -> `Descargar`.

Detail URLs are keyed on the billing close date: `/billing/detail/20260831`.

**The file count is not fixed.** MercadoLibre only offers report types that
have data for that period. Observed for BOCINA_SM on 2026-08-28:

| Period | Report types offered | Files |
|---|---|---|
| Agosto 2026 (EN CURSO) | BILL_ML, FULL, PAYMENT, NC_ML | 4 |
| Julio 2026 (cerrado) | BILL_ML, FULL, PAYMENT, NC_ML | 4 |
| Junio 2026 (cerrado) | BILL_ML, PAYMENT | 2 |

Default is the two newest cards including the EN CURSO month (Agosto + Julio).
Never hard-code an expected count - `--expect` defaults to 0 for this reason.

### Route 1 gotcha - the Reportes panel is in a shadow root

The reports UI is a micro-frontend rendered inside the shadow root of
`section.remote-module.fbi-billing-fe-reporting-reporting`. `querySelector`,
`By.XPATH` and `innerText` do not pierce shadow roots, so the panel looks
completely empty to ordinary Selenium lookups even while visible on screen.
All lookups go through the deep helpers in `meli_forms.py`, which recurse
through open shadow roots and return the element for a real Selenium click.

The select-all checkbox id is React-generated (e.g. `«r2»`) and changes between
renders - it must be matched by label text. The per-report ids are stable:
`BILL_ML`, `FULL`, `PAYMENT`, `NC_ML`.

### Route 1 gotcha - wait for the batch to settle

Select-all fires several downloads a few seconds apart. Returning at the first
completed file silently loses the rest (this produced a wrong "2 files per
month" reading during development). `wait_for_downloads()` instead waits until
no new file has appeared for `settle` seconds and nothing is still `.crdownload`.

### Route 2 - Excel de ventas

`Ventas` -> `/ventas/omni/listado` -> remove the applied **Envíos de hoy**
filter tag -> `Descargar Excel de ventas` -> the process-notification widget in
the corner -> `Descargar` on the newest entry.

Removing the filter is mandatory: with `Envíos de hoy` applied the list is
today's shipments only, often 0 ventas, and the Excel button stays `disabled`.
Generation is asynchronous - the widget row goes `Generando archivo` ->
`Archivo generado con éxito` (~25-30s) before its Descargar link works.

Yields exactly 1 file. Split across phases: `request_sales_excel()` presses the
button, `probe_sales_excel()` / `fetch_sales_excel()` collect it later.

The period dropdown is set to **Últimos 6 meses**, not the 2-month default —
otherwise older months are silently truncated out of the export. Match on the
label only: the option's full text carries live dates
(`Últimos 6 meses 28 feb. al 29 ago.`) that change daily.

### Route 3 - Reportes de stock (Full)

`Publicaciones` → `Control de stock` → the `Reportes` dropdown → 4 entries.
Yields 5 files (one report covers 2 periods). See FORMS.md reports 5–9.

⚠️ MercadoLibre changed this menu's markup mid-project: entries were
`li.andes-button-dropdown__menu-item` on 2026-08-31 and plain
`li.andes-list__item` on 2026-09-02, and the route reported `menu has: []` while
the menu was demonstrably open. Selectors accept either. When a route finds
nothing, check for a *changed class* before assuming a timing problem.

### Route 4 - MercadoPago reports

`/balance/reports/<kind>` — six report types sharing one UI in four flavours
(`typed`, `simple`, `accordion`, `wizard`). Driven by the `REPORTS` dict in
`mercadopago.py`; see FORMS.md reports 10–15.

⚠️ **The site's "today" is not this machine's today.** This box runs CST (UTC+8)
and MercadoPago México is UTC-6, so for much of the day the local clock is a day
ahead and an end date built from `date.today()` does not exist yet on the site.
`_select_range()` reads the calendar's own `--today` cell, clamps to it, and
returns the dates actually selected so callers rebuild their row-match string
from reality.

⚠️ **React re-renders invalidate a NodeList mid-loop.** Ticking several
checkboxes in one JS pass ticks only the first. Click one per call, from Python,
with a pause between.

A full run yields **18 files** on a good day: up to 8 billing + 1 ventas +
5 stock + 6 MercadoPago. Last verified end-to-end 2026-09-02 against
`BOCINA_SM` (6 billing — September offered only 2 report types — 1 ventas,
5 stock, 6 MercadoPago).

TODO - still to decide:

- How often the pull runs, and over which stores (see §10)
- Whether two stores can be open at once — this decides serial vs parallel

## 10. Running many stores

`run_batch.py` walks several stores in one client session.

```bash
.venv\Scripts\python -m platforms.mercadolibre.run_batch                       # every store in stores.batch
.venv\Scripts\python -m platforms.mercadolibre.run_batch --stores BOCINA_SM EWTTO_SM
.venv\Scripts\python -m platforms.mercadolibre.run_batch --only ventas --limit 2
.venv\Scripts\python -m platforms.mercadolibre.run_batch --dry-run             # list the stores, open nothing
```

It takes every `--only` / `--months` / `--collect-timeout` option
`run_downloads.py` takes, because both call the same `add_route_args()`.

### One pipeline, two entry points

`run_downloads.run_store(client, name, args, out_dir)` is the whole per-store
pipeline. `run_downloads.py` calls it once; `run_batch.py` calls it in a loop.
There is deliberately **no second copy** — a fix to a route reaches both.

### Which stores

**`stores.allowed` is a permission list, not a work list.** It says what the
tooling *may* open. Adding a second store there does not make
`run_downloads.py` download two — that script is single-store by design and
always runs exactly one (`stores.default`, or `--store`). Downloading several is
what `run_batch.py` is for. It now prints a note when other stores are allowed,
because expecting otherwise is the natural mistake.

`--stores` if given, else `stores.batch` from `config.json`, else the whole
`stores.allowed` list. Every name goes through `require_allowed()`, so a typo in
`--stores` is rejected locally without a request reaching Ziniao.

`batch` exists only for the case where the allowlist is deliberately wider than
the nightly run; leaving it `[]` (= all of `allowed`) is normal.

If `allowed` is empty — which means *unrestricted* — the batch **refuses to
start**. "Every store on the account" is 93 stores and is never what anyone
means.

### Serial, and why

One browser at a time. Whether Ziniao holds two stores open at once is still
untested, and each store carries its own proxy and fingerprint, so opening
several at once multiplies both load and the chance of tripping a verification
challenge. The client starts **once** and is reused, so the per-store cost is
`startBrowser`/`stopBrowser` — not a client restart.

`--pace` (default 5s) spaces the stores out.

### The batch does not wait for slow reports any more

A store's MercadoPago reports keep generating on MercadoPago's servers whether
or not its browser is open. Waiting for them inside the store's own session
therefore blocks every other store for nothing. Measured on the 13-store run of
2026-09-07: **14,424 s total, of which roughly 3,000 s was one store watching a
progress bar while twelve others sat idle** - TANKE_EE alone took 2,571 s
against a ~880 s baseline, and file count did not explain it (TRQWH_TA01 pulled
20 files in 881 s).

So a batch now runs in two passes:

```
store 1: request -> download the instant forms -> short collect -> defer the rest
store 2: same
   ...                                                    (registry written per store)
late collect: reopen store 1, 2, ... in the SAME order, short budget each
clean -> workbooks
```

Visiting the stores in the same order in the late pass is what gives the last
store its head start: it is collected last, so it gains the duration of every
other store's collection for free, without a third pass.

| Flag | Default | What it bounds |
|---|---|---|
| `--session-collect` | 300 s | the collect budget while the store is still open |
| `--late-timeout` | 420 s | the budget per store in the late pass |
| `--collect-timeout` | 1800 s | **single-store runs only** - `run_downloads.py` has no late pass, so it still waits properly |
| `--no-late-collect` | off | skip the late pass; everything owed stays in the registry |
| `--collect-only` | off | download nothing, just collect what is owed |

Measured over the same 12 stores, before and after:

| | per-store total | median store | **slowest store** |
|---|---|---|---|
| Before (2026-09-07) | 14,424 s | 897 s | **2,571 s** |
| After (2026-09-09) | 11,199 s | 997 s | **1,100 s** |

The headline is the last column: **no store stalls any more.** The long tail is
gone, which is what the change was for - 22% off the per-store total, and 57%
off the worst case. The median rose slightly (+100 s) because a store with one
straggler now spends its full `--session-collect` budget before deferring, where
a store with nothing outstanding used to leave immediately. Lower
`--session-collect` if that trade looks wrong for your store mix.

**Ventas is normally collected in-session**, because it is ready in 30-60 s and
because the store's accounting workbook needs it. But it *can* be deferred, and
is, if it misses the budget.

That was an open question until 2026-09-09, when it was tested directly:
requested in one browser session, the store closed, reopened on a different
debugging port, and the file collected on the first probe. Its readiness lives
in MercadoLibre's notification widget - server state, not anything the session
holds - which is the same reason a MercadoPago row survives a reopen.

Before that test ventas was excluded from deferral, which meant a ventas that
missed the budget was recorded as pending and then collected by nobody.

### logs/pending.json - what we have asked for but not collected

A request is a side effect on someone else's server. Once `Generar` is clicked
the report is being built whether or not this process survives, so the only way
to lose it is to forget we asked - which is exactly what happened to TANKE_EE
on 2026-09-07: it timed out with two reports still generating, nothing recorded
them, and the next run re-requested from scratch.

`pending_store.py` is that record. It is written **the moment each store
finishes**, not at the end of the batch, so a crash midway cannot lose the
reports already requested. An entry is removed only when its file is on disk, or
when the range turns out to hold no movements and will never produce one. A
probe that raises does **not** remove the entry — it used to, and one page
hiccup meant nobody ever went back for that report.

Each entry also carries what the last probe saw (`probes`, `first_probe_at`,
`last_probe_at`, `last_state`, `last_status_text`, `evidence`, `cause`), which
is what the batch's diagnosis reads.

An entry is not a handle into a browser session - it is the text of a table row
(report type, url, the exact period string), which is how `mercadopago.probe()`
finds it anyway. That is why a report survives the store being closed, the
process dying, or the machine rebooting.

To pick up what a previous run left generating, without downloading anything:

```bash
.venv\Scripts\python -m platforms.mercadolibre.run_batch --collect-only
```

It opens only the stores that actually owe something. Entries older than
`pending_store.MAX_AGE_DAYS` (7) are pruned at the start of every batch, since
an entry nobody has collected in a week is far more likely to be a stale record
than a live report. Pruned entries are not dropped silently: each is marked
`cause: EXPIRED`, appended to `logs/pending_expired.json`, and listed in the
batch output with its store, form, request time, probe count and last state.

### Failure isolation

Nothing one store does can end the batch.

- A **route** that breaks is caught inside `run_store()` by `_guard()`, recorded
  in `result["errors"]`, and the other three routes still run. Previously any
  route raising killed the whole invocation.
- A **store** that cannot be opened at all is caught by the batch loop, which
  moves to the next one. The `except` there is deliberately broad: the point of
  a batch is that store 7 dying in an unforeseen way still leaves 8–25 to run.
- `--stop-on-error` reverses this when you are debugging.

Isolating is not hiding — everything lands in the summary and the exit code:

| Code | Meaning |
|---|---|
| 0 | every store `ok` and nothing expected is missing |
| 1 | finished, needs attention: some store `partial` / `failed`, or reports still owed, or cleaning partially failed for this run's files |
| 2 | hard failure: Ziniao client error, or **no** store at all reached `ok` / `partial`, so nothing fresh was downloaded |
| 3 | a named store is not on the allowlist |
| 4 | `config.json` unreadable, or the batch list empty |

A store "needs attention" when it has route errors, still owes a report, or is
missing any expected report for a reason other than `NO_DATA` /
`NOT_AVAILABLE`. `run_downloads.py` follows the same rule for its one store: it
exits 1 on a route error or on a missing report that needs attention.

The last stdout line of a batch is `[BATCH_SUMMARY_JSON] <path>`; the runner
reads the summary from there and filters the line out of the human log.

### Why a report is missing

A run used to end with `尚未就绪` eleven cycles in a row and `仍欠 2 张` — no
reason, although at the moment it gave up the code usually *knew* the reason
(the row's state, a disabled button, a login page). `download_diag.py` keeps it:

- `expected_forms(args)` works out what this run **should** produce
  (`ventas`, `billing/2026-07` or `billing/1`…, `stock/…`, `mercadopago/<kind>`),
  so "missing" is computed rather than inferred from whichever warnings
  happened to print. A billing period that has not started yet is not expected.
- At the moment a report is given up on, `capture_evidence()` records the page
  address, title, `<html lang>`, visible banners, the row's status and a
  screenshot into `<download folder>/_diag/<HHMMSS>_<form>.json` + `.png`. It
  never raises and never waits.
- `classify()` — a pure function, no browser — turns that evidence into exactly
  one cause code from a closed set:

| Cause | Meaning | Whose problem | Retry helps? | What the operator does |
|---|---|---|---|---|
| `STORE_OPEN_FAILED` | The store browser did not open, or dropped mid-run | Ziniao | yes | Check the Ziniao client is logged in and the store opens by hand, then rerun the store |
| `LANGUAGE_WRONG` | The page is not in Spanish, so controls matched by Spanish text are not found | account | no | Set the account language back to Español in that store's browser, then rerun |
| `NO_PERMISSION` | This login has no permission to view reports | account | no | Ask the account admin to grant report permissions under Colaboradores |
| `PAGE_BLOCKED` | Not the expected page (login, verification or error page) | account | yes | Enter the store by hand, finish the login/verification, confirm the page opens, rerun |
| `PLATFORM_ERROR` | The platform returned a 5xx error page (temporary outage) | platform | yes | Rerun the store a little later; if it persists, check the platform's status or open the page by hand |
| `UI_NOT_FOUND` | Right page, but the control to click is missing or not clickable | script | no | Send the evidence screenshot to the maintainer: most likely a redesign, selectors need changing |
| `REQUEST_NOT_REGISTERED` | Generate was clicked, but the platform's list does not show the report | platform | yes | Rerun the store to request again; if it recurs, generate it once by hand on the page |
| `STILL_GENERATING` | The platform is still generating it | platform | yes | Run `run_batch.py --collect-only` later |
| `PLATFORM_DELAYED` | The platform says data is still processing, wait some hours | platform | yes | Run `run_batch.py --collect-only` in a few hours |
| `NO_DATA` | No data in the range; the platform will not produce a report | platform | no | Nothing |
| `NOT_AVAILABLE` | Not offered: period not started / no such period / store has no Full | platform | no | Nothing; if you believe it should exist, check the evidence |
| `DOWNLOAD_NOT_LANDED` | Report ready and download clicked, but no file landed | Ziniao | yes | Rerun to collect; if it recurs check disk space, download-folder permissions and the browser's download blocking |
| `FILE_REJECTED` | The file arrived but is unusable (empty, or no longer in the folder) | platform | yes | Download that report again |
| `INTERNAL_ERROR` | The script itself failed | script | yes | Send the exception to the maintainer; one rerun is reasonable |
| `EXPIRED` | Not collected within the retention period; no longer tracked | script | no | Request the report again if still needed |
| `UNKNOWN` | The evidence cannot tell the cause apart | script | yes | Open the evidence file and screenshot and judge by hand |

`NO_DATA` and `NOT_AVAILABLE` need nobody to do anything — the absence is
correct — so they do not make a store `partial` and do not move the exit code.
Every other cause does.

The diagnosis runs twice: once per store when it finishes, and once more for
the whole batch **after the late pass**, when a report that could only be
called `UNKNOWN` after one in-session probe usually has enough probes behind it
to be told apart. It only uses evidence already captured — no store is reopened
just to diagnose. The result is printed under `未下载的报表与原因` (one line per
store and cause, then a count by cause), and written to:

| Where | What |
|---|---|
| `REPORT.md` in the run folder | a `未下载 —— 原因` table: form, period, cause, retry?, next step, evidence file. Rewritten after the late pass, including for folders from an earlier run |
| `run_meta.json` in the run folder | `expected`, `missing` (the records, with evidence), `flags`, `cautions` |
| `logs/batch_<stamp>.json` | per-store `expected` / `missing` / `flags` / `cautions`, plus `missing_by_cause`, `stores_total` / `stores_ok` / `stores_attention` / `stores_failed`, `cleaning`, `expired`, `exit_code` |
| `logs/pending.json` | on each owed entry: `probes`, `absent_probes`, `first_probe_at`, `last_probe_at`, `last_state`, `last_status_text`, `evidence`, `cause` |
| `logs/pending_expired.json` | entries pruned after `MAX_AGE_DAYS`, marked `cause: EXPIRED` |

`flags` carries `ip_check_failed`, `range_clamped` and `locale`. `cautions` is
for downloads that succeeded but deserve a look — today only
`adopted_existing_row`: MercadoPago showed no new row after a request, so the
existing row was matched instead and the file may be an **earlier** report.
Both the requested and the adopted period are recorded, and the batch prints
them under `已下载但要留意`.

### Every run writes its own REPORT.md

`run_store()` drops `REPORT.md` into `downloads/<STORE>/<stamp>/` beside the
files, so the folder still explains itself when the terminal output is long
gone: run parameters, every file with its route and size, and separate sections
for `No data`, `Still owed` and `Errors`, plus a `未下载 —— 原因` section (the
table described under "Why a report is missing") with files collected in the
late pass, any `adopted_existing_row` cautions and run flags beneath it. The
batch replaces just that section after the late pass, since conclusions change
once owed reports arrive.

It is written for both entry points, since both go through `run_store()`. It
never raises — a run that landed 18 files must not be reported as failed because
a summary file could not be written.

Only the parameters that shaped the run are listed: an `--only ventas` run does
not print a billing period count, which would just make the reader wonder where
the billing files went.

### The run summary

Each batch writes `logs/batch_<stamp>.json` — per store: status, file count,
errors, what is still owed, what is missing and why, output folder, duration;
and for the batch: counts by cause, counts of stores by state, the cleaning
state and the exit code. A scheduled run throws stdout away, so without this a
silent failure is invisible until someone notices missing data.

Statuses are `ok`, `partial` (files landed, but a route errored **or an
expected report is missing or still owed**), `failed`, `not-allowed`. A store
with reports still owed used to be `ok`; so did a store whose billing or stock
route delivered nothing without raising.

MercadoPago has two ways of refusing to build a report, and both used to be
polled for the full 1800s because the row looks like one still generating:

- **no movements in the range** → `[no data]`, store stays `ok`. Nothing went
  wrong and that range will never produce a report.
- **"hay datos en proceso… espera unas horas"** → `[pending]`, marked
  `datos en proceso`. The request stands on MercadoPago's side; a later run
  collects it.

See FORMS.md, "Still generating and will never generate look identical".

⚠️ A route that quietly returns nothing counts as an **error**, not a success.
The `request_*` helpers warn and return `None` rather than raising so one bad
report cannot sink the other five — but on the first two-store run all six
MercadoPago requests returned `None` on EWTTO_SM and the store still reported
`ok, 13 files`. Silence is not success.

### Per-store permissions are a real failure mode

MercadoPago reports are gated per **collaborator login**, not per store.
EWTTO_SM's login lacks the permission, so every report page shows *"No tienes
permisos para ver los reportes"* and no `Crear reporte` button — which reads
exactly like a broken selector. `mercadopago._require_access()` detects the wall
before clicking anything and raises `NoReportAccess`, naming the two permissions
an account admin must grant under **Colaboradores**. Expect more of this as
stores are added: **verify each new store's MercadoLibre *and* MercadoPago
access before blaming the code.**

### The cleaning pipeline runs after the batch

`downloads/mx_sales` is a separate project that cleans the raw exports into
Parquet and DuckDB views — its own README, pyproject and tests. `run_batch.py`
runs it **once, after every store**, unless you pass `--no-transform`:

```
>>> transform: -m mx_sales build
    ventas_mx: 11498 rows from 23 file(s) -> …/data/processed/ventas_mx
    facturacion: 4138 rows from 31 file(s) -> …/data/processed/facturacion
    warehouse: …/data/mx_sales.duckdb (views: ventas_mx, facturacion)
    transform ok in 0 min 39 s
```

Three decisions behind that shape:

- **A subprocess, not an import.** A pandas or duckdb failure must not be able
  to reach back into a batch that just spent half an hour driving a browser. The
  worst a broken transform can do is print an error and set the exit code.
- **After all stores, not after each.** The pipeline is tree-shaped — its README
  says "a snapshot tree holds the same report many times over", and `discover`
  walks the whole raw directory. Once at the end sees every store; per store
  would repeat the same work N times.
- **The same interpreter.** It runs on `sys.executable`, so `pyarrow` and
  `duckdb` are pinned in our `requirements.txt` rather than living in a second
  environment.

If the stores all downloaded but the transform failed, the exit code is 1 and
the summary says which half broke — the files are on disk either way, and
`cd downloads && python -m mx_sales build` re-runs just the transform. A
cleaning failure is never counted as a store failure; it is reported in its own
`cleaning` field.

**Only this run's files count.** `build` walks the whole tree every time, so
one old file that cannot be cleaned used to make every later batch exit 1 — on
2026-10-09, 28 of 902 files failed, all of them September exports. The batch
now passes its run folders (this run's, plus any earlier folder the late pass
dropped a file into) as `transform.build(current_runs=…)` →
`build --current-run <folder>`. Failures elsewhere are printed as a `note:` and
returned as `failed_older`; if they are the only failures, the transform
counts as ok. What gets cleaned does not change — only how it is reported.

Those 28 files failed because the Ventas, Facturación and Notas de crédito
readers read the header from a fixed row, and MercadoLibre pushes the table
down a row when it puts a banner on top. The readers now locate the header by
column names (see `downloads/README.md`).

**Coverage is currently 2 of ~15 report types** (`ventas_mx`, `facturacion`).
The rest download fine and are simply not cleaned yet; `python -m mx_sales
discover --unmatched` lists them. Adding one is a module under
`downloads/mx_sales/reports/` plus a line in its `registry.py`.

### Still open

- **Parallel stores.** Get serial reliable across more than two first;
  concurrency multiplies load and detection risk.
- **Scheduling.** Ziniao is an Electron GUI driving a real Chromium window, so
  Task Scheduler's *"Run whether user is logged on or not"* (session 0, no
  interactive desktop) will not work. It must be *"Run only when user is logged
  on"*, on a machine that stays logged in and never sleeps.
- **Alerting.** Non-zero exit plus the summary JSON is the minimum; mailing the
  summary is the obvious follow-up.

---

## 10.5 The runner: `run_monthly.py`

Every production run goes through `run_monthly.py`. The name is historical —
registered scheduled tasks call `run_monthly.cmd`, and renaming it would break
them all.

```
月度报表.cmd    -> app.py         -> run_monthly.py --entrance manual ...
Task Scheduler (registered from 月度报表.cmd) -> run_monthly.cmd --entrance scheduled
by hand                            -> run_monthly.py        (--entrance defaults to cli)
```

Neither window spawns `run_batch.py` / `run_reports.py` itself any more. The
three paths differ in `--entrance` only, and that argument affects nothing but
the entrance column of the log. Before, the manual window chained the two
scripts itself and **omitted `--month` for the default month**, while the
scheduled task always passed it — and `run_batch.py` with and without `--month`
are two different download routes (Ventas window, how the billing page is
reached, the folder suffix). "Fine by hand, broken at night" could not be
reproduced from the window. The runner now **always** passes an explicit
`--month` to `run_batch.py`.

| Option | Meaning |
|---|---|
| `--month YYYY-MM` | target month, default the previous month |
| `--step both\|download\|report\|collect` | run one step only; `collect` = `run_batch.py --collect-only` |
| `--stores NAME …` | only these stores. Both children get `--stores`, and `run_reports.py` gets `--no-combined`: the roll-up's file name is fixed, so a two-store roll-up would overwrite the real one |
| `--out-dir DIR` | report output folder; otherwise the one chosen in the manual window (`gui_settings.json`), otherwise the default |
| `--entrance manual\|scheduled\|cli` | who started the run; default `cli` |
| `--max-hours N` | download deadline, default 5, `0` = none |
| `--dry-run` | print what would run |

What lives in the runner, once, for every entrance:

- **The lock** (`logs/monthly.lock`, created with `O_EXCL`). A second run exits
  5. A lock whose PID is gone is taken over at once; a lock whose holder has not
  sent a heartbeat for 30 minutes (the runner touches the lock with every
  child-process heartbeat) is taken over too — so a long `ALL` run is never
  taken over while it is alive, and a leftover lock after a power cut blocks
  for at most 30 minutes. `release()` only removes its own lock.
- **No report after a hard download failure.** Download exit 0 or 1 → reports
  are built (gaps show on sheet ⑩). Exit 2 / 3 / 4 → they are not: building
  then would produce a normal-looking workbook from an older folder.
- **Its own download deadline.** Task Scheduler's limit kills the whole chain —
  no report, no alert, no trace. At `--max-hours` the runner kills the download
  child and everything it started, **still builds the reports and writes the
  alert**, and exits 7. Task Scheduler's 6-hour limit remains only as a backstop
  for the runner itself hanging.
- **The alert file** `！报表异常_<YYYYMM>.txt` in the report folder (plus the
  optional webhook), for manual runs too. It is written when data will be
  missing from the reports or a report was not written: a hard download
  failure, stores failed / partial / still owed reports, the deadline, a report
  step exiting ≥ 2, or the runner itself crashing. A report step exiting 1
  (checks failed) does not alert, nor does a cleaning-only problem, an operator
  abort or a skipped run.

### Runs do not depend on a window

The manual window starts the runner as a detached process
(`runlog.spawn_detached`); closing the window does not stop the run. The viewer
reads the log **file**, not a pipe, so either window attaches to whatever run
is in progress — including a scheduled run that started on its own — and
reattaches after being reopened. While a run is in progress
(`logs/current.json` exists and its PID is alive) both windows lock their run
buttons, whoever started it: a second download would restart the Ziniao client
under the first.

Stopping is a file, not a signal: either window writes
`logs/runs/<run_id>.stop`, the runner checks for it every half second, kills
its child's process tree, records `aborted`, releases the lock and exits 6. If
the runner has not stopped within 20 s (`runlog.STOP_GRACE`), or was already
stalled, the window kills it by PID and finishes the bookkeeping for it.

A run whose heartbeat (`current.json` mtime, touched every 20 s) is older than
180 s while its process is still alive is shown as **not responding** rather
than as finished: it still holds the lock and Ziniao, so the buttons stay
locked and the stop button force-stops it.

### Run logs

```
logs/runs/<YYYYMMDD_HHMMSS>_<entrance>.log     the human log
logs/runs/<YYYYMMDD_HHMMSS>_<entrance>.json    summary: steps, exit codes, seconds, alert text, commands
logs/runs/<run_id>.stop                        only while a stop is requested
logs/current.json                              only while a run is in progress
```

`<entrance>` is `manual`, `scheduled` or `cli`. `logs/monthly_<YYYYMM>.log` /
`.json` are no longer written.

Every line has six columns separated by ` | `:

```
2026-10-13 14:05:01 | 定时 | 下载 | BOCINA_TA02 | 警告 | 没有出现新行……
time                | entrance | step | store or - | level | message
```

| Column | Values |
|---|---|
| entrance | `手动` (manual window) / `定时` (scheduled task, including "run now" and the rehearsal from the scheduling window) / `命令行` (typed, or `run_monthly.cmd` double-clicked) |
| step | `启动` / `下载` / `补收` / `清洗` / `报表` / `收尾` |
| store | the store name, or `-` |
| level | `信息` / `警告` / `错误` |

The entrance is on **every line**, not only in the header, because what gets
forwarded when something breaks is a dozen lines from the middle. The three
ways to tell the two entrances apart: that column, the file name, and the
coloured badge at the top-left of the viewer (also each entry of its run
picker).

The file opens with a header block (step `启动`): entrance, run id, month,
stores, steps, each command line, output folder, download deadline, machine and
user, Python, commit, log path.

The children's output is not rewritten. `runlog.LineClassifier` reads each line
of `run_batch.py` / `run_reports.py` and decides step, store and level from
the banners they already print; a line it cannot place is logged as `信息`
unchanged. The viewer filters by store and by level and can pin any past run.

A scheduled task registered **before** this change has no `--entrance
scheduled` in its action. It still runs the same process, but its lines say
`命令行`. Re-register once in the scheduling window ("设为每月自动跑"); the
window's status line says so until you do.

### Exit codes

`run_batch.py` is in §10. `run_reports.py`:

| Code | Meaning |
|---|---|
| 0 | every requested report written, all error-level checks pass, no source flagged incomplete by the platform |
| 1 | every requested report **written**, but something to look at: checks failed, formula errors, incomplete Ventas data, or a roll-up missing a store that had no usable folder |
| 2 | at least one requested report **not written** (crash, timeout, `--strict` abort, write failure, a store with no usable download folder), or no store usable at all |
| 3 | a store is not on the allowlist |
| 4 | config error |

The line between 1 and 2 is only whether the file was actually written this
time. Both used to be 1, so the runner could not tell "bridge B failed as it
does every month" from "there is no report".

`run_monthly.py` returns the worst of its children's 0–4, plus:

| Code | Meaning |
|---|---|
| 5 | another run holds the lock; this one was skipped |
| 6 | aborted by the operator |
| 7 | download deadline exceeded; reports were built from what had arrived |

A child killed by a signal or crashing with an out-of-range code is recorded
as 2.

### Report stage

- Each failed check is printed **once**, in its own report's `[校验]` block,
  with its troubleshooting hint directly beneath; the final summary adds one
  compact line per report (`未过 N 项：…` / `提醒 N 项：…` with the stores).
- A diagnostics file is written **only when a report was not written**, to
  `logs/report_diag/<report>.<stamp>.log`. The `<out-dir>/diagnostics/` folders
  of earlier versions are no longer produced and can be deleted.
- `--strict` now also blocks a Ventas export the platform itself flags as
  incomplete (`ventas_incomplete`, a warn-level check everywhere else). A
  strict abort leaves no file and counts as "not written" (exit 2).

---

## 11. Operational notes

- The client must be **fully exited** before relaunching in webdriver mode. This
  closes any stores a human has open — do not run this on a machine someone is
  actively working on.
- Leaving the client in webdriver mode is fine, but to return to normal use,
  exit from the tray and start it the usual way.
- Store caches live in `%LOCALAPPDATA%\SuperBrowser` and can grow large. The
  vendor demo has `delete_all_cache()` for this; only needed under disk pressure,
  and it fails while stores are open.
- Credentials belong in `config.json` only. It is gitignored — keep it that way,
  and rotate the password if it ever lands in a commit.
