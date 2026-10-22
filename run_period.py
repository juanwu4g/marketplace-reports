# -*- coding: utf-8 -*-
"""Period reports (quarterly, annual, custom) from the final monthly results:

    python run_period.py --quarter 2026-Q3 [--region MX] [--basis order] ...

The code lives in core/period_report/cli.py (usage and exit codes there;
docs/RESULTS_AND_PERIODS.md, section 5).
"""
import runpy

if __name__ == "__main__":
    runpy.run_module("core.period_report.cli", run_name="__main__", alter_sys=True)
