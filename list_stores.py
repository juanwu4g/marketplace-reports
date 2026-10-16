# -*- coding: utf-8 -*-
"""Kept at the repository root so existing commands keep working:

    python list_stores.py ...

The code lives in platforms/mercadolibre/list_stores.py (see docs/ARCHITECTURE.md).
"""
import runpy

if __name__ == "__main__":
    runpy.run_module("platforms.mercadolibre.list_stores", run_name="__main__", alter_sys=True)
