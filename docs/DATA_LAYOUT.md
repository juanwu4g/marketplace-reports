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
| anything else (incl. `_versions\`, `版本说明.txt`)          | `_旧文件\`                 |

Files already under `monthly\`, `quarterly\`, `annual\`, `custom\` or `_旧文件\`
are left alone, so `plan`/`copy` can run again. These are copies like every
other move: verified, listed in the manifest (which records the folder), the
flat originals deleted by `remove-old`; `results\**.json` `workbook` paths into
that folder are rewritten to the new place. `undo` leaves the sorted copies
where they are (the flat originals are still there until `remove-old`, which is
where the old layout looks). The setting itself is kept; the tool says so after
switching config.json. If the shared drive cannot be opened, the tool refuses
to run until it can (or the setting is changed).

## The cleaning pipeline

The Mexican cleaning project lives in `pipelines\` (code). It is always told its
data folders through `MX_SALES_*` (`platforms/mercadolibre/transform.data_env`):
in the new layout it reads `raw\ML\MX\` and writes to `work\mx_sales\`. By hand:
`python -m platforms.mercadolibre.transform build`.
