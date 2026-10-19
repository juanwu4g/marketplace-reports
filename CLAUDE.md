# CLAUDE.md — context for Claude Code sessions in this repo

## What this is

A system that downloads seller reports from marketplace back offices through the
Ziniao anti-detect browser and builds monthly financial workbooks (one per store,
one combined per country). MercadoLibre is done for Mexico, Chile and Peru; Falabella for Chile
(`platforms/falabella/`, market `FA-CL`). The
goal is to add more platforms (TikTok Shop, Temu, Falabella, Walmart, Amazon,
SHEIN …) to this same repo, each as `platforms/<name>/`, sharing `core/` and the
one window, with "sister" reports that look like one family.

The owner is not a developer: explain in plain words, give step-by-step commands
for git and deployment, and say which folder/branch you mean. Operators and
finance read Chinese; all UI text, logs and workbook labels are Simplified
Chinese. Developer docs are English.

## Read before changing anything

- `docs/ARCHITECTURE.md` — layout, run flow, invariants, extension points (§5 lists
  what must change when platform #2 arrives)
- `docs/NEW_PLATFORM_SOP.md` — the procedure for a new platform
- `docs/REPORT_FAMILY.md` — what every platform's workbook must look like
- `docs/PLATFORM_PLAYBOOK.zh.md` — lessons from MercadoLibre that hold on any platform
- `platforms/mercadolibre/README.md`, `FORMS.md` — the reference implementation

## Hard rules

- **Never open a store unless it is meant for this test.** The production machine
  runs the monthly job for real stores; opening a store someone or something is
  using breaks their session. Discovery and test runs happen in the Windows test
  VM (below), whose `config.json` allowlist holds only the stores being tested.
  Always pass explicit `--stores`; never start a batch without a store list.
- **Never commit `config.json` or anything under `downloads/` or `logs/`** —
  credentials and customer personal data. Tests use synthetic fixtures.
- **Numbers must not move silently.** `tests/test_golden.py` rebuilds real
  store-months (data outside git in `~/report-golden/`, private) cell by cell. Only
  rebaseline after the owner/finance accept a change. Accounting rule changes are
  the owner's decision — ask, don't decide.
- `run_monthly.cmd` stays at the root (scheduled tasks call it). `.cmd` files are
  ASCII only (Chinese messages go in `launcher/*.txt`).
- Don't push; the owner pushes (this box has no GitHub credentials). Commit only
  when asked or at clear checkpoints, as `JW <jw2021grad@gmail.com>` with a
  `Co-Authored-By: Claude …` trailer, on a feature branch, never on `main`.

## Commands

```
# tests (Ubuntu dev box)
.venv/bin/python -m pytest tests -q
cd downloads && ../.venv/bin/python -m pytest tests -q
.venv/bin/python -m pytest tests/test_golden.py -rs   # see which golden cases ran

# Windows / production
.venv\Scripts\python run_monthly.py --region CL --stores "<store>" --month 2026-09
.venv\Scripts\python run_monthly.py --region FA-CL --stores "<store>" --month 2026-09   # Falabella Chile
.venv\Scripts\python list_stores.py --mx|--cl|--pe       # lists, opens nothing
.venv\Scripts\python -m tools.probe_pages --store <name> --out <dir> <url> …
.venv\Scripts\python -m tools.probe_session --store <name> --cmd-dir <dir> --out <dir>

# new platform (dev box): scaffold, then register in core/markets.py by hand
.venv/bin/python -m tools.new_platform AZ amazon "亚马逊" MX
.venv/bin/python -m pytest tests/test_platform_contract.py -q
```

## Where things are (this machine)

- Branches: `main` (old), `unify-entrances` (Mexico refactor), `multi-region`
  (+ Chile, Peru, one window), `platform-core` (this folder: core/ + platforms/
  layout + docs). Each newer branch contains the older ones. Other folders on this
  box are worktrees of the same repo: `~/meli-mx-pipeline` (main),
  `~/meli-mx-unify`, `~/meli-ml-multi` (multi-region).
- Windows test VM: `~/vm/win11/` (start `run-vm.sh`; SSH `-p 2222` with
  `id_ed25519`; VNC 127.0.0.1:5901). Test folder for this branch
  `C:\Users\meli\marketplace-reports`; copy code with `~/vm/win11/syncp`, run on the VM
  desktop with `~/vm/win11/pljob <name> <script> args…` (log in
  `logs\job_<name>.log`, ends with `EXIT=<code>`). Ziniao cannot open windows from
  the SSH session — use the job helpers.
- Golden data: `~/report-golden/{mx,cl,pe}/<store>/` + `<case>.baseline.json`.
