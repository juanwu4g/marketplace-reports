# Where the data lives

**Audience:** developers and whoever looks after the production machine.

Code and data are separate. The code folder (`D:\marketplace-reports`, git) holds
only code, `config.json`, `gui_settings.json` and the double-click `.cmd` files.
All data — downloads, reports, results, logs — lives under one folder named in
`config.json`:

```json
"paths": {"data_root": "D:/marketplace-data"}
```

`core/layout.py` is the only place that knows the folders; every module asks it.

## New layout (data_root set)

```
D:\marketplace-data\
├── reports\                                 what people open
│   ├── monthly\2026-09\
│   │   ├── ML-MX\                           final workbooks of the month, both ML bases,
│   │   │   ├── BOCINA_TA02_2026-09_账单基准.xlsx      and the country summaries
│   │   │   ├── 版本说明.txt                  which version is final and why
│   │   │   └── _versions\<version>\         every generated version (cleaned once final)
│   │   ├── ML-CL\  ML-PE\  FA-CL\  AZ-MX\
│   │   ├── 本月总览_2026-09.xlsx            every store's grade, headline figures, versions, 需要你处理
│   │   └── ！需要处理_202609.txt            only while something needs a person (one per month)
│   ├── quarterly\2026-Q3\ML-MX\MX_ML_2026-Q3_账单基准.xlsx
│   ├── annual\2026\…
│   └── custom\2026-07_2026-11\…
├── raw\<PF>\<CC>\<store>\<YYYY-MM>\<run>\    downloads (files + run_meta.json + REPORT.md)
├── results\<PF>\<CC>\<store>\<YYYY-MM>\      monthly results, one file per version
├── work\                                     intermediate files (cleaning pipeline output)
└── logs\
```

`<run>` is the download run's name (`20261021_233002_m202609`); it is also the
first half of a report version's id, so a report can be traced to its download.
The window's 「报表输出到」 setting moves `reports\` (e.g. to a shared drive).

## Renamed stores

A store renamed in the Ziniao browser (2026-10: `清货2_EE02` → `QH_EE02`) keeps its
old folders under `raw\` and `results\`. Nothing on disk needs moving: write the
new name everywhere in `config.json` (`allowed`, `batches`) and put the old
name(s) in the store's `stores.catalog` entry under `formerly` (a list, or one
string):

```json
"catalog": {
  "QH_EE02": {"platform": "ML", "country": "MX", "formerly": ["清货2_EE02"]}
}
```

From then on every reader (`core/results.py`, `core/layout.run_folders`) treats
the old name's folders as the store's: a month's versions come from both
folders and the final is the newest countable one across them; the period
reports and 本月总览 show the store once, under the new name, and log
`旧名 清货2_EE02 的结果并入 QH_EE02`; a report built right after the rename still
finds the newest download even when it sits under the old name. New downloads
and results are always written under the current name; the clean-up never
deletes a folder under an old name. The old folders may be archived later, once
no month still needs them (a period report reads them for as long as they hold
a month's final). A former name must not still be a current store, and one old
name belongs to one store: a wrong entry makes only that store 配置有问题
(`config_problems`), the others run as usual. `formerly` exists only in the
`catalog` form of the config (the old `regions` form has no place for it).

## Old layout (no data_root)

Everything where it has always been: downloads in `downloads\<store>\<run>\`,
reports flat in `downloads\data\reports\financial\`, logs in `logs\`, results in
`downloads\data\results\`. A machine stays on it until it is moved.

## Moving a machine to the new layout

`tools/migrate_layout.py` (see its docstring): lists what it would move, copies
everything into the new tree, verifies every file, writes `paths.data_root` into
config.json, and only on a second, explicit command deletes the old folders.
Run it between monthly runs with the scheduled task paused.

When the window's 「报表输出到」 names its own folder (the production machine:
`S:\Finance\电商数据` on the shared drive), that folder is the flat report
folder in the old layout and the reports root in the new one, so the tool
sorts its files **in place**, under that same folder, and never under
`data_root\reports`:

| in `S:\Finance\电商数据` (old, flat)                         | after (same folder)        |
|-------------------------------------------------------------|----------------------------|
| `BOCINA_TA02_2026-09.xlsx`, `…_账单基准.xlsx`, `…_指标.json`, `MX_ML_全店汇总_2026-09.xlsx`, `全店_2026-09_指标.json` | `monthly\2026-09\ML-MX\` |
| `！需要处理_202609.txt`, `！报表异常_*_202609.txt`, `本月总览_2026-09.xlsx` | `monthly\2026-09\` |
| `期间报表\quarterly\2026-Q3\ML-MX\*` (same for `annual`, `custom`) | `quarterly\2026-Q3\ML-MX\*` |
| `_versions\`, `版本说明.txt` (ours, old versions)            | `_旧文件\`                 |
| **anything else** (the finance team's own folders and files, e.g. `10月\MercadoLibre\…`) | **left exactly where it is**, counted in the plan |

That folder is shared with people, so only what this program wrote is touched
(the production drive held 1,536 of the finance team's files next to our 108 on
2026-11-04). Excel's `~$…` lock files are never files to copy. Files already
under `monthly\`, `quarterly\`, `annual\`, `custom\` or `_旧文件\` are left
alone, so `plan`/`copy` can run again. These are copies like every
other move: verified, listed in the manifest (which records the folder), the
flat originals deleted by `remove-old`; `results\**.json` `workbook` paths into
that folder are rewritten to the new place. `undo` leaves the sorted copies
where they are (the flat originals are still there until `remove-old`, which is
where the old layout looks). The setting itself is kept; the tool says so after
switching config.json. If the shared drive cannot be opened, the tool refuses
to run until it can (or the setting is changed). A file that cannot be read
during `copy` (open in Excel, a share hiccup) does not stop the copy: it is
listed at the end with its reason, the manifest records it under `unreadable`,
config.json is **not** switched, and running `copy` again after closing the file
finishes the job (identical copies are skipped).

## The cleaning pipeline

The Mexican cleaning project lives in `pipelines\` (code). It is always told its
data folders through `MX_SALES_*` (`platforms/mercadolibre/transform.data_env`):
in the new layout it reads `raw\ML\MX\` and writes to `work\mx_sales\`. By hand:
`python -m platforms.mercadolibre.transform build`.
