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
│   │   └── ！报表异常_*.txt
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

## The cleaning pipeline

The Mexican cleaning project lives in `pipelines\` (code). It is always told its
data folders through `MX_SALES_*` (`platforms/mercadolibre/transform.data_env`):
in the new layout it reads `raw\ML\MX\` and writes to `work\mx_sales\`. By hand:
`python -m platforms.mercadolibre.transform build`.
