# -*- coding: utf-8 -*-
"""
Amazon downloads for several stores in one Ziniao session.

    python -m platforms.amazon.run_batch --month YYYY-MM --market AZ-MX --stores "<store>"
    python -m platforms.amazon.run_batch --collect-only --month YYYY-MM --stores ...
    python -m platforms.amazon.run_batch --dry-run --market AZ-MX

The batch itself (store choice, allowlist and platform checks, exit codes 0–4,
logs/batch_AZ_<stamp>.json, logs/pending_AZ.json, the owed catch-up, the
`[BATCH_SUMMARY_JSON]` last line) is shared: core/platform_batch.py. This module
only says what is Amazon's own.

Amazon's own: its routes (forms, run_downloads — whose require_implemented() is
checked before Ziniao starts) and site facts.
"""
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import platform_batch
from platforms.amazon import forms
from platforms.amazon import periods
from platforms.amazon import run_downloads
from platforms.amazon import site

PLATFORM = "AZ"
DEFAULT_MARKET = "AZ-MX"

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
