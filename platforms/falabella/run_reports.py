# -*- coding: utf-8 -*-
"""
Falabella monthly financial reports: one workbook per store + one combined.

    python -m platforms.falabella.run_reports --month 2026-09
    python -m platforms.falabella.run_reports --month 2026-09 --stores "DAB_CL-Fala (Coco)"
    python -m platforms.falabella.run_reports --market FA-CL --no-per-store
    python -m platforms.falabella.run_reports --dry-run

The runner itself (CLI, exit codes 0–4, workbook names, the lines core/runlog
reads) is shared: core/platform_reports.py. This module only says what is
Falabella's own.

Falabella's own: each store uses its newest download folder **for that month**
(downloads/<store>/<stamp>_m<YYYYMM>/) that holds a Transacciones file
(NewReportTransaction_*.xlsx); without one the store is left out.
"""
from core import console  # noqa: F401  中文输出编码保护
from core import platform_reports
from platforms.falabella import periods

PLATFORM = "FA"
DEFAULT_MARKET = "FA-CL"


def folder_markers(month):
    return ("NewReportTransaction_*.xlsx",)


RUNNER = platform_reports.ReportRunner(
    PLATFORM, __package__, DEFAULT_MARKET, periods,
    folder_markers=folder_markers,
    no_folder_reason="没有 %s 的下载目录（或目录里没有交易明细）")

# the runner's steps, by their old names (tests call them)
latest_folder = RUNNER.latest_folder
resolve = RUNNER.resolve
safe_build = RUNNER.safe_build
exit_code = platform_reports.exit_code
main = RUNNER.main


if __name__ == "__main__":
    platform_reports.run_as_script(main)
