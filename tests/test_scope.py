# -*- coding: utf-8 -*-
"""core/scope.py：平台 → 国家 → 店铺 这棵树，和勾选之后跑什么。"""
import json
import os

from core import markets
from core import runlog
from core import scope as SC

GROUPS = [("MX", ["BOCINA_SM", "EWTTO_SM"]), ("CL", ["Tecnohub_CL", "Andes_CL"]),
          ("PE", []), ("FA-CL", ["AIS_CL-Fala", "DAB_CL-Fala"]), ("AZ-MX", ["SM-Amazon"])]


def t():
    return SC.tree(GROUPS)


def test_tree_is_platform_first_from_markets_and_config():
    tree = t()
    assert [p.code for p in tree] == ["ML", "FA", "AZ"]
    assert [p.name for p in tree] == ["美客多 MercadoLibre", "Falabella", "亚马逊 Amazon"]
    assert [u.unit for u in tree[0].units] == ["MX", "CL"]        # PE has no store: hidden
    assert tree[1].units[0].name == "智利" and tree[2].units[0].stores == ["SM-Amazon"]
    assert SC.unit_label(tree[1].units[0]) == "Falabella · 智利"


def test_a_new_country_slots_in_by_the_markets_order(monkeypatch):
    """Adding a country is markets.py + config only: it appears in its place."""
    ml = markets.PLATFORMS["ML"]
    monkeypatch.setitem(markets.PLATFORMS, "ML", ml._replace(countries=ml.countries + ("CO",)))
    tree = SC.tree([("CO", ["Bogota_CO"])] + GROUPS)              # config order does not matter
    assert [u.unit for u in tree[0].units] == ["MX", "CL", "CO"]
    assert tree[0].units[2].name == "哥伦比亚"


def test_region_arg_all_list_and_subset():
    tree = t()
    sel = SC.everything(tree)
    assert SC.region_arg(sel, tree) == "ALL" and SC.subset(sel, tree) is None
    del sel["CL"]
    assert SC.region_arg(sel, tree) == "MX,FA-CL,AZ-MX"
    sel = {"FA-CL": ["DAB_CL-Fala"], "MX": None}
    assert SC.region_arg(sel, tree) == "MX,FA-CL"                  # tree order
    assert SC.subset(sel, tree) == ["DAB_CL-Fala"]
    assert SC.region_arg({"FA-CL": ["AIS_CL-Fala", "DAB_CL-Fala"]}, tree) == "FA-CL"
    assert SC.region_arg({}, tree) is None
    # every store of every unit listed one by one is still ALL
    assert SC.region_arg({u.unit: list(u.stores) for u in SC.units(tree)}, tree) == "ALL"


def test_normalize_drops_unknown_and_empty():
    tree = t()
    sel = SC.normalize({"XX": None, "CL": ["Gone_CL"], "MX": ["EWTTO_SM", "Nope"]}, tree)
    assert dict(sel) == {"MX": ["EWTTO_SM"]}


def test_scope_text_and_groups():
    tree = t()
    sel = {"MX": None, "CL": ["Andes_CL"], "AZ-MX": None}
    assert SC.scope_text(sel, tree) == "美客多：墨西哥 2 家、智利 1/2 家；亚马逊：墨西哥 1 家 —— 共 4 家"
    assert [(u.unit, ss, part) for u, ss, part in SC.groups(sel, tree)] == [
        ("MX", ["BOCINA_SM", "EWTTO_SM"], False), ("CL", ["Andes_CL"], True),
        ("AZ-MX", ["SM-Amazon"], False)]
    assert SC.scope_text({}, tree) == "一个都没勾"


def test_from_region_arg_and_settings_round_trip():
    tree = t()
    assert SC.from_region_arg("ALL", tree) == SC.everything(tree)
    assert dict(SC.from_region_arg("FA-CL,MX", tree, ["AIS_CL-Fala"])) == {
        "MX": None, "FA-CL": ["AIS_CL-Fala"]}
    sel = {"CL": ["Andes_CL"], "FA-CL": None}
    saved = json.loads(json.dumps(SC.to_settings(sel, tree)))
    assert dict(SC.from_settings(saved, tree)) == {"CL": ["Andes_CL"], "FA-CL": None}
    assert SC.from_settings({"units": {"XX": None}}, tree) == SC.everything(tree)
    assert SC.from_settings(None, tree, default={"MX": None}) == {"MX": None}


def _run(log_dir, rid, doc):
    p = runlog.run_paths(rid, log_dir)
    os.makedirs(os.path.dirname(p["log"]), exist_ok=True)
    open(p["log"], "w").close()
    runlog.write_json(p["json"], dict(doc, run_id=rid))


def test_last_results_per_unit_newest_wins(tmp_path):
    d = str(tmp_path / "logs")
    _run(d, "20261015_020000_scheduled",
         {"region": "ALL", "exit": 1, "status": "done",
          "regions": [{"region": "MX", "exit": 0, "status": "done"},
                      {"region": "FA-CL", "exit": 1, "status": "done"}]})
    _run(d, "20261019_100000_manual",
         {"region": "MX,FA-CL", "exit": 0, "status": "done",
          "regions": [{"region": "FA-CL", "exit": 0, "status": "done",
                       "all_stores": False}]})
    _run(d, "20261020_100000_manual", {"region": "AZ-MX", "exit": 2, "status": "failed",
                                        "stores": "all"})
    res = SC.last_results(d)
    assert SC.unit_status(res, "MX") == "10-15 ✔ 正常"
    assert SC.unit_status(res, "FA-CL") == "10-19 ✔ 正常（部分店）"
    assert SC.unit_status(res, "AZ-MX") == "10-20 ✗ 出错"
    assert SC.unit_status(res, "CL") == ""
    assert SC.last_results(str(tmp_path / "none")) == {}
