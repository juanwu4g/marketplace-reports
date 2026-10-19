# -*- coding: utf-8 -*-
"""
Falabella downloads for several stores in one Ziniao session.

    python -m platforms.falabella.run_batch --month 2026-09 --stores "DAB_CL-Fala (Coco)"
    python -m platforms.falabella.run_batch --market FA-CL --month 2026-09
    python -m platforms.falabella.run_batch --collect-only --month 2026-09 --stores ...
    python -m platforms.falabella.run_batch --dry-run --market FA-CL

The batch itself (store choice, allowlist and platform checks, exit codes 0–4,
logs/batch_FA_<stamp>.json, logs/pending_FA.json, the owed catch-up, the
`[BATCH_SUMMARY_JSON]` last line) is shared: core/platform_batch.py. This module
only says what is Falabella's own.

Falabella's own: its routes (forms, run_downloads), site facts, and the catch-up
prints the files of every owed report it collected (log_collected).
"""
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import platform_batch
from platforms.falabella import forms
from platforms.falabella import periods
from platforms.falabella import run_downloads
from platforms.falabella import site

PLATFORM = "FA"
DEFAULT_MARKET = "FA-CL"

RUNNER = platform_batch.BatchRunner(
    PLATFORM, DEFAULT_MARKET,
    forms=forms, periods=periods, run_downloads=run_downloads, site=site,
    log_collected=True)

# the batch's steps, by their old names (tests and tools call them)
resolve_stores = RUNNER.resolve_stores
batch_verdict = RUNNER.batch_verdict
row_of = RUNNER.row_of
collect_owed = RUNNER.collect_owed
store_needs_attention = platform_batch.store_needs_attention
write_summary = platform_batch.write_summary
failed_row = platform_batch.failed_row
main = RUNNER.main


if __name__ == "__main__":
    sys.exit(main())
