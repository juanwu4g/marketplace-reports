# -*- coding: utf-8 -*-
"""Kept at the repository root so existing commands keep working:

    python open_store.py ...

The code lives in platforms/mercadolibre/open_store.py (see docs/ARCHITECTURE.md).
"""
import runpy

if __name__ == "__main__":
    runpy.run_module("platforms.mercadolibre.open_store", run_name="__main__", alter_sys=True)
