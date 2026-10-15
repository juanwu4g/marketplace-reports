# -*- coding: utf-8 -*-
"""Kept at the repository root so existing commands keep working:

    python run_monthly.py ...

The code lives in core/run_monthly.py (see docs/ARCHITECTURE.md).
"""
import runpy

if __name__ == "__main__":
    runpy.run_module("core.run_monthly", run_name="__main__", alter_sys=True)
