# -*- coding: utf-8 -*-
"""
Liverpool downloads for several stores in one Ziniao session.

    python -m platforms.liverpool.run_batch --month YYYY-MM --market LV-MX --stores "<store>"
    python -m platforms.liverpool.run_batch --collect-only --month YYYY-MM --stores ...
    python -m platforms.liverpool.run_batch --dry-run --market LV-MX

The batch itself (store choice, allowlist and platform checks, exit codes 0–4,
logs/batch_LV_<stamp>.json, logs/pending_LV.json, the owed catch-up,
the `[BATCH_SUMMARY_JSON]` last line) is shared by every platform:
core/platform_batch.py. This module only says what is this platform's own —
usually nothing to change here. The hooks (log_collected, name_zh, …) are listed
in core/platform_batch.py; a fix to the batch goes there, once, for everyone.
"""
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import platform_batch
from platforms.liverpool import forms
from platforms.liverpool import periods
from platforms.liverpool import run_downloads
from platforms.liverpool import site

PLATFORM = "LV"
DEFAULT_MARKET = "LV-MX"

RUNNER = platform_batch.BatchRunner(
    PLATFORM, DEFAULT_MARKET,
    forms=forms, periods=periods, run_downloads=run_downloads, site=site)

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
