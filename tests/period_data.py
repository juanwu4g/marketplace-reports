# -*- coding: utf-8 -*-
"""Synthetic monthly results for the period-report tests (no real store data).

configure(tmp_path, monkeypatch, v2=True) points config.json (and, for the old
layout, the code root) at a temporary folder; doc(...) makes a schema-1 results
document; put(...) writes one with core/results.write.
"""
import json

from core import paths
from core import results as R
from core import store_config

CATALOG = {"S1": {"platform": "ML", "country": "MX"}, "S2": {"platform": "ML", "country": "MX"},
           "S3": {"platform": "ML", "country": "MX"}, "BOCINA_SM": {"platform": "ML", "country": "MX"},
           "F1": {"platform": "FA", "country": "CL"}, "F2": {"platform": "FA", "country": "CL"},
           "A1": {"platform": "AZ", "country": "MX"}}
CURRENCY = {"MX": ("MXN", 2), "CL": ("CLP", 0)}


def configure(tmp_path, monkeypatch, v2=True, catalog=None):
    """A config.json with the catalog; v2 adds paths.data_root. Returns the
    reports root a run would use (no --out-dir)."""
    cat = dict(CATALOG if catalog is None else catalog)
    cfg = {"stores": {"allowed": list(cat), "catalog": cat}}
    if v2:
        cfg["paths"] = {"data_root": str(tmp_path / "data")}
    else:
        monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    p = tmp_path / "config.json"
    p.write_text(json.dumps(cfg), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(p))
    from core import run_monthly
    monkeypatch.setattr(run_monthly, "GUI_SETTINGS", str(tmp_path / "gui_settings.json"))
    from core import layout
    return layout.reports_root()


def doc(store, month, built, level="complete", pl=None, basis=None, reasons=(), checks=(), sku=None,
        inventory=None, fees=None, returns=None, memo=None):
    """One results document. built: 'YYYYMMDD_HHMMSS'. pl: [(key, value)] or
    [(key, value, kind)]; labels are the keys in upper case."""
    pf, cc = CATALOG[store]["platform"], CATALOG[store]["country"]
    cur, dec = CURRENCY[cc]
    basis = basis or ("bill" if pf == "ML" else "main")
    run = "%s_000000_m%s" % (built[:8], month.replace("-", ""))
    rows = []
    for item in pl if pl is not None else [("gmv", 100.0), ("fee", -10.0), ("net", 90.0, "f")]:
        key, value = item[0], item[1]
        kind = item[2] if len(item) > 2 else "d"
        rows.append({"key": key, "label": key.upper(), "original": "orig " + key, "kind": kind, "value": value})
    return {"schema": 1, "platform": pf, "market": "%s-%s" % (pf, cc), "country": cc, "currency": cur,
            "decimals": dec, "store": store, "month": month, "basis": basis, "run": run,
            "built_at": "%s-%s-%sT%s:%s:%s" % (built[:4], built[4:6], built[6:8], built[9:11], built[11:13],
                                                built[13:15]),
            "version": R.version_id(run, built), "source_folders": [], "workbook": None,
            "code": {"commit": "test", "fingerprint": "x"},
            "grade": {"level": level, "reasons": list(reasons)}, "checks": list(checks),
            "missing": [], "flags": {"provisional": [], "data_incomplete": False},
            "pl": rows, "memo": memo or [], "fees": fees or [], "returns": returns or [],
            "sku": sku or [], "inventory": inventory}


def put(*args, **kw):
    d = doc(*args, **kw)
    R.write(d)
    return d
