# -*- coding: utf-8 -*-
"""更早几趟留下的欠账不能算到这一趟头上。

2026-10-13 实测（定时入口，BOCINA_TA02）：这一趟 20 张报表全部下到，却因为
上一趟留在登记簿里的 3 条旧欠账被标成 partial、退出码 1，还写了告警文件。
"""
from core import download_diag
from core import pending_store
from platforms.mercadolibre import run_batch

NOW = "C:/dl/BOCINA_TA02/20261013_165113_m202609"
OLD = "C:/dl/BOCINA_TA02/20261013_161937_m202609"
EXPECTED = ["ventas", "mercadopago/settlement", "mercadopago/after_collection",
            "mercadopago/account_statement", "billing/2026-09"]


def _old_entry(report):
    return {"kind": "mercadopago", "report": report, "store": "BOCINA_TA02",
            "out_dir": OLD, "range": "14/ago/2026 a 12/oct/2026"}


def _rec(form, cause, owed=False):
    return {"form": form, "period": None, "cause": cause, "owed": owed,
            "explain": download_diag.CAUSES[cause][0]}


def _patch_remove(monkeypatch):
    removed = []
    monkeypatch.setattr(pending_store, "remove", lambda es: removed.extend(es))
    return removed


def test_form_delivered_now_cancels_the_old_owed_entry(monkeypatch, capsys):
    removed = _patch_remove(monkeypatch)
    row = {"store": "BOCINA_TA02", "status": "ok", "out_dir": NOW,
           "expected": EXPECTED, "pending": ["something from the old folder"]}
    live = [_old_entry("after_collection"), _old_entry("settlement")]
    carried = [dict(_rec("mercadopago/after_collection", "REQUEST_NOT_REGISTERED", True),
                    carried_from="20261013_161937_m202609"),
               dict(_rec("mercadopago/settlement", "PLATFORM_DELAYED", True),
                    carried_from="20261013_161937_m202609")]
    recs = []                                  # 这一趟全部到齐
    run_batch.settle_carried(row, recs, carried, live, NOW, EXPECTED)
    assert removed == live, "这一趟下到了，旧条目就该撤掉"
    assert all(r["superseded"] for r in carried)
    assert row["pending"] == []
    row["missing"], row["carried"] = recs, carried
    assert not run_batch.store_needs_attention(row), "旧欠账不影响这一趟的状态"
    assert "撤销更早目录里的 2 条旧欠账" in capsys.readouterr().out


def test_form_still_missing_now_keeps_the_old_entry_but_counts_once(monkeypatch):
    removed = _patch_remove(monkeypatch)
    row = {"store": "BOCINA_TA02", "status": "ok", "out_dir": NOW,
           "expected": EXPECTED, "pending": []}
    live = [_old_entry("settlement")]
    carried = [dict(_rec("mercadopago/settlement", "PLATFORM_DELAYED", True),
                    carried_from="20261013_161937_m202609")]
    recs = [_rec("mercadopago/settlement", "PLATFORM_DELAYED", True)]
    run_batch.settle_carried(row, recs, carried, live, NOW, EXPECTED)
    assert removed == [], "这一趟也没下到：旧条目留着，之后还能补取"
    assert not carried[0]["superseded"]
    assert len(row["pending"]) == 1 and "PLATFORM_DELAYED" in row["pending"][0]
    row["missing"], row["carried"] = recs, carried
    assert run_batch.store_needs_attention(row)
    v = run_batch.batch_verdict([row])
    assert v["missing_by_cause"] == {"PLATFORM_DELAYED": 1}, "同一张不能数两遍"


def test_carried_records_are_listed_apart_from_this_run(monkeypatch):
    row = {"store": "BOCINA_TA02", "missing": [],
           "carried": [dict(_rec("mercadopago/after_collection",
                                 "REQUEST_NOT_REGISTERED", True),
                            carried_from="20261013_161937_m202609",
                            superseded=True)]}
    text = "\n".join(download_diag.format_section([row]))
    assert "没有。应下的报表全部到齐。" in text
    assert "更早几轮留下的（不计入本轮" in text
    assert "来自 20261013_161937_m202609 | 最新一趟已下到，旧条目已撤销" in text
    assert text.index("应下的报表全部到齐") < text.index("更早几轮留下的")


def test_collect_only_measures_against_the_newest_run_on_disk(tmp_path, monkeypatch):
    """--collect-only 没有自己的目录。2026-10-13 实测：更早一趟里一条申请失败的
    记录（平台 503）和两条旧欠账，在后来那一趟把报表全部下到之后，仍被补取那
    一轮当成"这家店缺报表"报出来，还写了告警。"""
    import json, os
    store = tmp_path / "BOCINA_TA02"
    old = store / "20261013_161937_m202609"
    new = store / "20261013_165113_m202609"
    for d, missing in ((old, [_rec("mercadopago/account_statement", "PLATFORM_ERROR"),
                              _rec("mercadopago/after_collection",
                                   "REQUEST_NOT_REGISTERED", True)]),
                       (new, [])):
        d.mkdir(parents=True)
        (d / "run_meta.json").write_text(json.dumps(
            {"expected": EXPECTED, "missing": missing}), encoding="utf-8")
    assert run_batch.newest_run([str(old)]) == (str(new), EXPECTED)

    entry = dict(_old_entry("after_collection"), out_dir=str(old))
    monkeypatch.setattr(pending_store, "load", lambda: [entry])
    removed = _patch_remove(monkeypatch)
    row = {"store": "BOCINA_TA02", "status": "ok", "out_dir": None,
           "pending": ["after_collection ..."], "errors": []}
    run_batch.finalize_diagnosis([row])
    assert row["missing"] == [], "最新一趟什么都不缺"
    assert removed == [entry]
    assert {r["form"] for r in row["carried"]} == {
        "mercadopago/account_statement", "mercadopago/after_collection"}
    assert all(r["superseded"] for r in row["carried"])
    assert row["status"] == "ok" and row["pending"] == []
    assert run_batch.batch_verdict([row])["exit_code"] == 0
