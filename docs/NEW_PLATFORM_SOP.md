# SOP — adding a new platform

**Audience:** the business owner who knows where the reports are, the finance
person who signs off the numbers, and the developer (human or Claude session)
who builds the automation. Each step says who does it.

**Outcome:** a new folder `platforms/<platform>/` that downloads every form a store
needs for a month, explains every form it could not get, and builds a monthly
workbook per store plus a combined one — in the same family layout as the
MercadoLibre reports — started from the same window and the same scheduled task.

Read first: `ARCHITECTURE.md` (how the system is laid out) and
`PLATFORM_PLAYBOOK.zh.md` (what we learned the hard way; it holds on any platform).

---

## 0. Before starting: decide and write down (owner + finance, ~1 hour)

- [ ] **Which stores** (exact Ziniao names; `python list_stores.py` lists platforms
      and stores without opening any), which countries, which month to test with.
- [ ] **What the report must answer.** Usually: revenue, platform fees by type,
      refunds, what actually reached the bank, margin before product cost, and
      "does the money tie out". Anything platform-specific (warehouse fees,
      ads, penalties).
- [ ] **Is there an official seller API?** (Amazon SP-API, TikTok Shop Open API,
      Walmart Marketplace API …). If yes and the company can get access, prefer it
      over browser clicking for the forms it covers: APIs change far less often
      than web pages. Use Ziniao only for what the API does not give.
- [ ] **A test machine.** Discovery and first runs never happen on the production
      machine: Ziniao is shared there and the monthly job may be using it. Use the
      VM or a second Windows machine with its own `config.json` whose allowlist
      contains only the new platform's test stores.

## 1. Record where the forms are (owner, per form)

Fill in `FORM_DISCOVERY_TEMPLATE.md` once per form — one copy per form, in
`platforms/<platform>/discovery/`. Screenshots or a short screen recording of
the clicks are worth more than prose. The essentials:

- the page (URL or menu path) and the buttons, in order
- filters and date range used, and whether the platform limits the range
- synchronous download, or "we'll prepare it" (where does the finished file
  appear; email? a list? a bell?)
- one real exported file for one month

## 2. Discovery in the browser (developer, on the test machine)

1. Add the test stores to the test machine's `config.json` allowlist (only those).
2. **Probe, don't guess.** Two tools, both refusing any store not on the
   allowlist:
   - `python -m tools.probe_pages --store <name> --out <dir> <url> …` — a quick
     look: opens the store, visits each page and saves a screenshot plus the
     visible text (and, with `--js file.js`, whatever a small script returns —
     e.g. the buttons and their labels), then closes the store.
   - `python -m tools.probe_session --store <name> --cmd-dir <dir> --out <dir>
     [--max-minutes 120]` — a conversation with the page: opens the store **once**,
     sends downloads to `<out>/files`, writes `READY` into the command folder and
     then runs every new `<name>.py` dropped there, writing `<name>.out` (what it
     printed, or the traceback) ending with `<<DONE>>`. Scripts see `d` (the
     driver), every helper of `core/browser_kit.py` (`find`, `need`, `js_click`,
     `real_click`, `table_rows`, `wait_new_files` …) and `dump(d, "tag")`
     (screenshot + JSON with URL, title, text including shadow roots, and the
     controls with their attributes). Variables persist from one script to the
     next. Drop a file `STOP` to end it (it also ends after `--max-minutes`).
     On the VM, start it with the job helper (Ziniao cannot open windows from an
     SSH session), then copy command files in and read the `.out` files back.
   Read the dump, then write the next step. Keep the one-off command scripts out
   of the repo; what they teach goes into `FORMS.md` and the routes.
3. Confirm per form: how to reach it, how to set the period, how to know it is
   ready, how to download it, what the file looks like (header row, sheet name,
   column names in the store's language).
4. Measure the numbers in playbook §8 (generation time, retention, range limits).
5. Write what you found into `platforms/<platform>/FORMS.md` (same structure as
   `platforms/mercadolibre/FORMS.md`).

**Safety rules (non-negotiable):** never open a store not on the test allowlist;
pass explicit store names to every batch; never `driver.quit()` (it kills the
Ziniao browser); close stores through the Ziniao client.

## 3. Build the download layer (developer)

Start from the scaffold — it already keeps the shared contract:

```
python -m tools.new_platform AZ amazon "亚马逊" MX      # CODE PACKAGE 中文名 COUNTRY
```

It creates `platforms/<package>/` and prints the line to add to
`core/markets.py` `PLATFORMS` (add it by hand, then run
`python -m pytest tests/test_platform_contract.py -q`). `--dry-run` works at
once; a real run says "还没实现" until the pieces below are written:

| File | Job | Model it on |
|---|---|---|
| `site.py` | per-country facts: hosts, locale, currency, timezone, quirks | `falabella/site.py`, `mercadolibre/meli_site.py` |
| `periods.py` | `form_ids(month)`: the forms a store should yield for a month; `in_report(form)`: which ones the report reads | `falabella/periods.py` |
| `forms.py` (one or more) | the routes: request / probe readiness / fetch, each giving up with a reason (`ctx.give_up`). Browser helpers from `core/browser_kit.py` (shadow roots, `need`, `real_click`, `wait_new_files`, `retried`); put only this platform's page quirks in a `page_kit.py` next to it | `falabella/forms.py`, `falabella/page_kit.py` |
| `run_downloads.py` | one store end to end; writes `run_meta.json` and `REPORT.md` (scaffolded; set `IMPLEMENTED = True` when the routes exist) | `falabella/run_downloads.py` |
| `run_batch.py` | many stores in one session; owed catch-up; summary JSON for the runner (scaffolded, usually unchanged) | `falabella/run_batch.py` |

Rules that carry over unchanged (details in the playbook):
- two phases: request everything slow first, download fast things meanwhile,
  collect at the end; anything still generating is "owed" and collected later
- verify the date range actually applied (read it back from the page)
- every missing form gets a cause code and evidence (screenshot + page facts)
- language gate: check the UI language before downloading; column names follow it
- ASCII-only `.cmd`; child processes started with `-m package.module`

Done when: one test store yields every form for the test month, twice in a row,
and the batch summary lists nothing unexplained.

## 4. Understand the money (developer + finance) — before writing the report

Do this in Excel first, by hand, for one store and one month:
- map every fee/charge name to a category (commission, shipping, ads, storage /
  fulfilment, penalties, other, tax withheld)
- find what proves the money: the platform's payout/settlement file vs. its order
  list. Which keys link them (order id, payment id)?
- tax: VAT rate, whether the platform withholds anything, what kind of invoice
  the platform issues for its fees
- refunds, cancellations, partial refunds: how they appear in each file
- write the agreed rules down in `platforms/<platform>/README.md` ("Accounting
  rules" section) and get finance to confirm them

Code implements rules that are already agreed. It is not the place to discover them.

## 5. Build the report (developer)

Follow `REPORT_FAMILY.md`: same sheet order, same styles, same check levels, same
⑩ "gaps and to-dos" sheet, same combined workbook naming
(`<PLATFORM>_<COUNTRY>_全店汇总_<month>.xlsx`). Platform-specific content goes in
the sheets' bodies, not in a different layout.

- `run_reports.py` (scaffolded) + `report_build.py` in the platform folder. The
  platform writes load / derive / validate; the sheets come from
  `core/report_family.py` (`REPORT_FAMILY.md` §6): a P&L row spec with
  formulas as data, bridge lines with a residual row, column specs for the long
  tables, the ⑩ gaps sheet. Don't copy layout code from another platform —
  if a sheet needs something the family lacks, add it to `report_family` with a
  test, so every platform gets it
- every check is `error` (numbers can't be trusted) or `warn` (look at this);
  each one says in plain Chinese what it compared and what to do
- tests with **synthetic** fixtures (`tests/test_<platform>_*.py`). Real exports
  contain customer personal data and must never be committed.
- add a golden case once finance has accepted one real month: copy that store's
  download folder to `~/report-golden/<country>/<store>/`, add it to `CASES` in
  `tests/test_golden.py`, run `python tests/test_golden.py --rebaseline <case>`

## 6. Plug into the runner and the window (developer)

- register the platform (code, package, Chinese name, Ziniao platform-name pattern,
  countries) in `core/markets.py` — `tools/new_platform.py` printed the line —
  and its stores in `config.json` (`allowed`, `catalog` with platform + country,
  `batches["<CODE>-<COUNTRY>"]`)
- nothing to change in `core/run_monthly.py` or the window: they run every
  configured unit; the platform's `run_batch` / `run_reports` only have to keep the
  shared contract (ARCHITECTURE §5: `--month`, `--stores`, `--market`,
  `--collect-only`, `--dry-run`, exit codes, `[BATCH_SUMMARY_JSON]`).
  `tests/test_platform_contract.py` checks it for every registered platform
  whose package imports — it must pass before the platform goes near a test
  machine
- run the whole chain from the window on the test machine; then from a
  registered scheduled task (try-run button)

## 7. Accept and deploy (finance + owner)

- [ ] finance reviews one real month per country and signs off the open questions
- [ ] all tests pass: `python -m pytest tests -q` and the cleaning package's tests
- [ ] the golden tests of every existing platform still pass unchanged
- [ ] merge to `main`; on production: `git pull`, update `config.json`, run the new
      platform once from the window, then let the scheduled run include it
- [ ] update `README.md` (operators) with anything they must do differently

## 8. Definition of done for a platform

1. Every store of the platform downloads every needed form for a month, from the
   scheduled task, unattended — or the window says exactly which store needs a
   person and why.
2. The per-store and combined workbooks open with no formula errors, the bridges
   tie out or explain the difference, and ⑩ lists what is still missing.
3. `FORMS.md`, `README.md` (platform notes + accounting rules) and the discovery
   records exist in `platforms/<platform>/`.
4. Tests (synthetic) and one golden test exist; all existing tests still pass.

---

## Appendix — rough effort, from the MercadoLibre experience

Downloads for a new platform are usually the smaller part; agreeing the
accounting rules and making every check tie out takes longer. Plan for one
discovery session with the owner per platform, a first working download within
days, and several rounds with finance before the numbers are signed off. A new
**country** of an existing platform (Chile, Peru) took about a day each,
including real test runs.
