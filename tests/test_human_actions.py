# -*- coding: utf-8 -*-
"""要人去店铺里处理的事：看得见、能一键补跑。

2026-10-14 那轮 12 家店，UNIT_PW01(catalog) 每条路线都撞上 MercadoLibre 的
登录验证页，一张报表都没下到；这件事只埋在 2,300 行日志中间。
"""
import json
import os

from core import download_diag
from core import run_monthly
from core import runlog

from test_runner import env, child, plan, _summary  # noqa: F401  复用夹具


def _rec(form, cause):
    return {"form": form, "cause": cause, "period": None,
            "explain": download_diag.CAUSES[cause][0], "evidence": {}}


ROWS = [
    {"store": "UNIT_PW01(catalog)", "status": "partial",
     "missing": [_rec("ventas", "PAGE_BLOCKED"),
                 _rec("mercadopago/settlement", "PAGE_BLOCKED")]},
    {"store": "EWTTO_SM", "status": "partial",
     "missing": [_rec("mercadopago/settlement", "NO_PERMISSION"),
                 _rec("billing/2026-09/BILL_MP_V3", "UI_NOT_FOUND")]},
    {"store": "TANKE_EE", "status": "partial",
     "missing": [_rec("mercadopago/after_collection", "PLATFORM_ERROR")],
     "carried": [_rec("ventas", "PAGE_BLOCKED")]},
    {"store": "BOCINA_TA02", "status": "ok", "missing": []},
]


def test_only_causes_a_person_must_fix_are_listed():
    acts = download_diag.human_actions(ROWS)
    assert [(a["store"], a["cause"]) for a in acts] == [
        ("UNIT_PW01(catalog)", "PAGE_BLOCKED"), ("EWTTO_SM", "NO_PERMISSION")]
    login = acts[0]
    assert login["title"] == "要登录 / 验证" and login["rerun"] is True
    assert login["forms"] == ["ventas", "mercadopago/settlement"], "同店同因并成一条"
    assert acts[1]["rerun"] is False, "缺权限：开通之前重跑没用"
    # 平台出错、脚本问题、更早几轮留下的都不算"要人处理"
    assert not any(a["store"] == "TANKE_EE" for a in acts)


def _write_run(log_dir, run_id, stores, rows, month="2026-09"):
    os.makedirs(runlog.runs_dir(log_dir), exist_ok=True)
    batch = os.path.join(log_dir, "batch_%s.json" % run_id)
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": rows}, f, ensure_ascii=False)
    runlog.write_json(runlog.run_paths(run_id, log_dir)["json"],
                      {"run_id": run_id, "month": month, "stores": stores,
                       "status": "done", "exit": 1, "batch_summary": batch})
    open(runlog.run_paths(run_id, log_dir)["log"], "w").close()


def test_rerunning_one_store_clears_it_and_keeps_the_others(tmp_path):
    d = str(tmp_path / "logs")
    _write_run(d, "20261014_025744_manual", "all", ROWS)
    got = runlog.pending_human_actions(d)
    assert got["month"] == "2026-09"
    assert {a["store"] for a in got["actions"]} == {"UNIT_PW01(catalog)", "EWTTO_SM"}
    # 登录好以后只重跑了 UNIT_PW01：它以重跑那轮为准（好了），EWTTO_SM 仍以全店那轮为准
    _write_run(d, "20261014_120000_manual", ["UNIT_PW01(catalog)"],
               [{"store": "UNIT_PW01(catalog)", "status": "ok", "missing": []}])
    got = runlog.pending_human_actions(d)
    assert [a["store"] for a in got["actions"]] == ["EWTTO_SM"]
    assert got["actions"][0]["run_id"] == "20261014_025744_manual"


def test_a_report_only_run_does_not_hide_the_list(tmp_path):
    d = str(tmp_path / "logs")
    _write_run(d, "20261014_025744_manual", "all", ROWS)
    rid = "20261014_130000_manual"            # 只出报表：没有批次汇总
    runlog.write_json(runlog.run_paths(rid, d)["json"],
                      {"run_id": rid, "month": "2026-09", "stores": "all",
                       "status": "done", "exit": 1, "batch_summary": None})
    open(runlog.run_paths(rid, d)["log"], "w").close()
    assert len(runlog.pending_human_actions(d)["actions"]) == 2


def test_rerun_downloads_the_stores_but_reports_for_everyone():
    p = run_monthly.build_plan("both", "2026-09", ["UNIT_PW01(catalog)"], "/out",
                               py="py", reports_all=True)
    dl, rp = p[0]["cmd"], p[1]["cmd"]
    assert dl[-2:] == ["--stores", "UNIT_PW01(catalog)"]
    assert "--stores" not in rp and "--no-combined" not in rp, "全店汇总要重出"
    assert "全部店铺" in p[1]["label"]
    # 默认（子集）行为不变
    q = run_monthly.build_plan("both", "2026-09", ["X"], "/out", py="py")
    assert "--no-combined" in q[1]["cmd"]


def test_runner_argv_carries_reports_for():
    a = runlog.runner_argv("manual", "both", "2026-09", ["X"], None, reports_for="all")
    assert a[-2:] == ["--reports-for", "all"]
    assert "--reports-for" not in runlog.runner_argv("manual")


def test_run_logs_lists_and_alerts_on_stores_needing_a_person(env, tmp_path):  # noqa: F811
    s = _summary(tmp_path, stores=ROWS)
    code = env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert code == 1
    rid, doc, recs = env.last()
    assert [a["store"] for a in doc["human_actions"]] == ["UNIT_PW01(catalog)", "EWTTO_SM"]
    errs = [r for r in recs if r["level"] == "错误"]
    assert any("需要人工处理的店铺" in r["message"] for r in errs)
    assert any(r["store"] == "UNIT_PW01(catalog)" and "登录" in r["message"] for r in errs)
    text = env.alert_text()
    assert "2 家店铺需要人工处理" in text and "重跑这些店铺并重出全部报表" in text
    assert text.index("需要人工处理") < text.index("部分报表没下成"), "放在告警最前面"


def test_exit_2_says_which_step_failed():
    t = runlog.exit2_text({"steps": [{"step": "download", "code": 1},
                                     {"step": "report", "code": 2}]})
    assert t.startswith("下载完成，但有报表没生成") and "紫鸟" not in t
    t = runlog.exit2_text({"steps": [{"step": "download", "code": 2}]})
    assert t.startswith("下载失败")
    assert runlog.exit2_text({"steps": []}) is None
