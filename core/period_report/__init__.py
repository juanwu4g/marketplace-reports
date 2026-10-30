# -*- coding: utf-8 -*-
"""Period reports (quarterly, annual, custom) built from final monthly results.

docs/RESULTS_AND_PERIODS.md §5 is the contract. Input is only
core/results.final() per store-month — never platform data. One workbook per
market (currencies differ) with every store of the market that has results in
the period; the report family's ten sheets (core/report_kit, core/report_family).

    model.py     periods, gathering the finals, every sum (to the cent); which
                 store-months are missing and which are before the store's
                 first data month (未上线 — not a gap, never an alert)
    workbook.py  the ten sheets + the hidden version record
    output.py    generate(): names, folders, previous file, alerts, automatic periods
    choice.py    the window tab's rules (no tkinter)
    cli.py       run_period.py (repository root)

Callers: run_period.py, core/run_monthly (after a run's report step, when the
month closes a quarter / the year), core/app (the 期间报表 tab).
"""
from core.period_report.output import (  # noqa: F401
    COMPLETE, EMPTY, ERROR, INCOMPLETE, STATUS_CN, alert_name, auto_periods, clear_alert, generate,
    markets_of, summary_lines, target, workbook_name, write_alert)
from core.period_report.model import (  # noqa: F401
    Period, basis_cn, basis_of, custom, gather, period_cn, quarter, year)
