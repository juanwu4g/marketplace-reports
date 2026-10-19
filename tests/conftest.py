# -*- coding: utf-8 -*-
"""Tests never read the machine's real config.json.

It holds credentials and that machine's store list, and it differs between the
dev box, the VM and production: a test that passes only because of what one
machine's file happens to contain is not a test. Every test starts from a small
neutral old-format config (no catalog — every store is MercadoLibre Mexico unless
`regions` says otherwise); tests that need something specific stub
store_config._stores themselves, as before.
"""
import json

import pytest

NEUTRAL = {"stores": {"default": "BOCINA_SM", "allowed": [], "batch": []}}


@pytest.fixture(autouse=True)
def _neutral_config(tmp_path_factory, monkeypatch):
    from core import store_config
    p = tmp_path_factory.getbasetemp() / "neutral_config.json"
    if not p.exists():
        p.write_text(json.dumps(NEUTRAL), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(p))
