# -*- coding: utf-8 -*-
"""
Liverpool monthly financial reports: one workbook per store + one combined.

    python -m platforms.liverpool.run_reports --month YYYY-MM --market LV-MX
    python -m platforms.liverpool.run_reports --month YYYY-MM --stores "<store>" --no-combined
    python -m platforms.liverpool.run_reports --dry-run

The runner itself (CLI, exit codes 0–4, workbook names, the lines core/runlog
reads) is shared by every platform: core/platform_reports.py. This module only
says what is this platform's own.

Liverpool's own: each store uses its newest download folder **for that month**
(core/layout.run_folders_for_month) that holds the month's ledger
(LV_mayor_<YYYYMM>.xlsx — every figure of the report comes from it), else the newest
one with a run_meta.json (a failed run is then reported as such by the report's
checks). Files missing there are looked for in the month's other folders by
report_build.
"""
from core import console  # noqa: F401  中文输出编码保护
from core import platform_reports
from platforms.liverpool import periods
from platforms.liverpool import readers

PLATFORM = "LV"
DEFAULT_MARKET = "LV-MX"


def folder_markers(month):
    return (readers.LEDGER % periods.stamp(month), "run_meta.json")


RUNNER = platform_reports.ReportRunner(
    PLATFORM, __package__, DEFAULT_MARKET, periods,
    folder_markers=folder_markers)

# the runner's steps, by their old names (tests call them)
latest_folder = RUNNER.latest_folder
resolve = RUNNER.resolve
safe_build = RUNNER.safe_build
exit_code = platform_reports.exit_code
main = RUNNER.main


if __name__ == "__main__":
    platform_reports.run_as_script(main)
