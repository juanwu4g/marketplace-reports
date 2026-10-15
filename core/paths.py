# -*- coding: utf-8 -*-
"""Where things live on disk. Every module asks here instead of looking at its own __file__.

ROOT is the repository root: config.json, downloads/, logs/, launcher/ and the
double-click .cmd files are there. Modules moved into core/ and platforms/<name>/,
so "the folder this file is in" is no longer the root.
"""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
