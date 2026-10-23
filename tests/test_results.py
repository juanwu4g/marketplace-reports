# -*- coding: utf-8 -*-
"""core/results.py: file names, validation, the final-version rule, period months."""
import json
import os

import pytest

from core import results as R
from core import store_config


@pytest.fixture
def cfg(tmp_path, monkeypatch):
    c = tmp_path / "config.json"
    c.write_text(json.dumps({"paths": {"data_root": str(tmp_path / "data")},
                             "stores": {"allowed": ["S1"], "catalog": {"S1": {"platform": "ML", "country": "MX"}}}}),
                 encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(c))
    return tmp_path


def doc(run, built, level, basis="bill", month="2026-09"):
    return {"schema": 1, "platform": "ML", "market": "ML-MX", "country": "MX", "currency": "MXN",
            "decimals": 2, "store": "S1", "month": month, "basis": basis, "run": run,
            "built_at": "%s-%s-%sT%s:%s:%s" % (built[:4], built[4:6], built[6:8], built[9:11], built[11:13],
                                                built[13:15]),
            "version": R.version_id(run, built), "grade": {"level": level, "reasons": []}, "checks": [],
            "pl": [{"key": "gmv", "label": "GMV", "original": "", "kind": "d", "value": 1.0}]}


def test_write_read_and_final_rule(cfg):
    R.write(doc("20261003_000000_m202609", "20261003_040000", R.COMPLETE))
    R.write(doc("20261012_000000_m202609", "20261012_040000", R.GAPS))          # newer but not countable
    R.write(doc("20261012_000000_m202609", "20261012_050000", R.COMPLETE, basis="order"))
    vs = R.versions("S1", "2026-09", "bill")
    assert [v["grade"]["level"] for v in vs] == [R.GAPS, R.COMPLETE]           # newest first
    assert R.final("S1", "2026-09", "bill")["run"] == "20261003_000000_m202609"
    R.write(doc("20261026_000000_m202609", "20261026_040000", R.COMPLETE_WARN))
    assert R.final("S1", "2026-09", "bill")["run"] == "20261026_000000_m202609"   # newest countable wins
    assert R.final("S1", "2026-08", "bill") is None
    assert R.stores() == ["S1"] and R.months("S1") == ["2026-09"]


def test_malformed_documents_are_refused_and_listed(cfg):
    bad = doc("20261003_000000_m202609", "20261003_040000", "great")
    with pytest.raises(ValueError):
        R.write(bad)
    d = os.path.dirname(R.path_for("S1", "2026-09", "x__20261003_040000", "bill"))
    os.makedirs(d)
    open(os.path.join(d, "junk.json"), "w").write("{")
    open(os.path.join(d, "x__20261003_040000.bill.json"), "w").write("{")
    assert R.versions("S1", "2026-09") == []
    assert len(R.problems("S1", "2026-09")) == 2


def test_period_months():
    assert R.quarter_months("2026-Q3") == ["2026-07", "2026-08", "2026-09"]
    assert R.quarter_months("2026q4") == ["2026-10", "2026-11", "2026-12"]
    assert R.months_between("2026-11", "2027-02") == ["2026-11", "2026-12", "2027-01", "2027-02"]
    assert len(R.year_months(2026)) == 12
    assert R.quarter_of("2026-09") == "2026-Q3" and R.quarter_of("2026-10") == "2026-Q4"
    with pytest.raises(ValueError):
        R.months_between("2026-09", "2026-08")
    assert R.official_basis("ML") == "bill" and R.official_basis("FA") == "main"
