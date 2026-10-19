# -*- coding: utf-8 -*-
"""bootstrap.py, the self-check: requirements pins are all checked, and the
store setup is validated the way the runner reads it — without opening anything."""
import json

import pytest

import bootstrap
from core import store_config

GOOD = {"default": "BOCINA_SM",
        "allowed": ["BOCINA_SM", "DAB_CL-Fala (Coco)", "SM-Amazon"],
        "catalog": {"BOCINA_SM": {"platform": "ML", "country": "MX"},
                    "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"},
                    "SM-Amazon": {"platform": "AZ", "country": "MX"}},
        "batches": {"ML-MX": ["BOCINA_SM"], "FA-CL": ["DAB_CL-Fala (Coco)"],
                    "AZ-MX": ["SM-Amazon"]}}


@pytest.fixture
def fresh(monkeypatch):
    monkeypatch.setattr(bootstrap, "problems", [])

    def use(cfg):
        d = json.loads(json.dumps(cfg))
        monkeypatch.setattr(store_config, "_stores", lambda: d)
        return d
    return use


def test_every_pinned_requirement_is_listed():
    names = [n for n, _ in bootstrap.requirements()]
    for want in ("requests", "selenium", "pandas", "openpyxl", "pypdf"):
        assert want in names


def test_good_store_setup_passes(fresh, capsys):
    fresh(GOOD)
    bootstrap.check_stores({"stores": GOOD})
    out = capsys.readouterr().out
    assert bootstrap.problems == []
    assert "亚马逊 墨西哥：1 家店" in out and "Falabella 智利：1 家店" in out


def test_store_missing_from_catalog_fails(fresh):
    d = fresh(GOOD)
    d["allowed"].append("Mystery")
    bootstrap.check_stores({"stores": d})
    assert bootstrap.problems and "店铺配置有误" in bootstrap.problems[0][0]


def test_batch_with_another_markets_store_fails(fresh):
    d = fresh(GOOD)
    d["batches"]["AZ-MX"] = ["SM-Amazon", "BOCINA_SM"]
    bootstrap.check_stores({"stores": d})
    assert bootstrap.problems and "亚马逊 墨西哥" in bootstrap.problems[0][0]


def test_empty_allowlist_fails(fresh):
    fresh({"default": "X", "allowed": []})
    bootstrap.check_stores({"stores": {}})
    assert bootstrap.problems and "allowed" in bootstrap.problems[0][0]
