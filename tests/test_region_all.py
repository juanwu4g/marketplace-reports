# -*- coding: utf-8 -*-
"""--region ALL：全部国家依次跑，一把锁、一份日志，各国各出各的报表和告警。

子进程一律换成几行 Python；config.json 用桩。绝不起真正的 run_batch。
"""
import datetime
import io
import json
import os
import sys
import threading
import time

import pytest

from core import run_monthly
from core import runlog
from core import store_config

PY = sys.executable
MONTH = "2026-09"
MX = ["BOCINA_SM", "EWTTO_SM"]
CL = ["Tecnohub_CL", "Andes_CL"]

CFG = {"default": "BOCINA_SM",
       "allowed": MX + CL,
       "batch": MX,
       "regions": {"Tecnohub_CL": "CL", "Andes_CL": "CL"},
       "batches": {"CL": CL}}


@pytest.fixture
def cfg(monkeypatch):
    d = json.loads(json.dumps(CFG))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


@pytest.fixture
def env(tmp_path, monkeypatch, cfg):
    logs = tmp_path / "logs"
    out = tmp_path / "out"
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(logs))
    monkeypatch.setattr(run_monthly, "LOCK", str(logs / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    monkeypatch.setattr(run_monthly, "known_stores", lambda: [])
    monkeypatch.setattr(run_monthly, "POLL", 0.1)

    class Env(object):
        log_dir = str(logs)
        out_dir = str(out)

        def args(self, argv):
            return run_monthly.parse_args(
                ["--month", MONTH, "--out-dir", self.out_dir] + argv)

        def run(self, argv, plans=None):
            return run_monthly.execute(self.args(argv), plan=plans)

        def last(self):
            rid = runlog.list_runs(self.log_dir)[0]["run_id"]
            p = runlog.run_paths(rid, self.log_dir)
            with io.open(p["log"], encoding="utf-8") as f:
                recs = [runlog.parse_line(x) for x in f.read().splitlines()]
            assert all(recs), "每一行都必须是标准的六列"
            return rid, runlog.read_json(p["json"]), recs

        def alert(self, region):
            return os.path.join(self.out_dir, run_monthly.alert_name(MONTH, region))

    return Env()


def child(code, text=""):
    return [PY, "-c", "import sys\nsys.stdout.reconfigure(encoding='utf-8')\n"
                      "print(%r, flush=True)\nsys.exit(%d)\n" % (text, code)]


def plan(*steps):
    return [{"key": k, "label": k, "cmd": c} for k, c in steps]


# ------------------------------------------------------------ 命令

def test_region_all_parses():
    assert run_monthly.parse_args(["--region", "all"]).region == "ALL"
    assert run_monthly.parse_args(["--region", "ALL"]).region == "ALL"


def _dry(env, capsys, argv):
    assert env.run(argv + ["--dry-run"]) == 0
    lines = [l.strip() for l in capsys.readouterr().out.splitlines()
             if "run_batch" in l or "run_reports" in l]
    assert lines, "dry run printed no child command"
    return lines


@pytest.mark.parametrize("step", ["both", "download", "report", "collect"])
def test_each_region_gets_exactly_the_commands_it_gets_alone(env, capsys, step):
    """墨西哥的命令一字不差；智利的和单独跑智利时一样（显式 --stores）。"""
    extra = [] if step == "both" else ["--step", step]
    mx = _dry(env, capsys, extra)
    cl = _dry(env, capsys, extra + ["--region", "CL"])
    both = _dry(env, capsys, extra + ["--region", "ALL"])
    assert both == mx + cl
    assert all("--region" not in l and "--stores" not in l for l in mx)
    if step != "report":
        assert all("--stores Tecnohub_CL Andes_CL" in l
                   for l in cl if "run_batch.py" in l)


def test_dry_run_lists_regions_in_order_and_skips_the_empty_one(env, capsys):
    assert env.run(["--region", "ALL", "--dry-run"]) == 0
    out = capsys.readouterr().out
    i_mx = out.index("第 1/2 个国家：墨西哥（MX），2 家店铺")
    i_cl = out.index("第 2/2 个国家：智利（CL），2 家店铺")
    assert i_mx < i_cl
    assert "秘鲁（PE）：config.json 里没有店铺，跳过" in out


def test_store_subset_is_refused_with_all(env, capsys):
    assert env.run(["--region", "ALL", "--stores", "BOCINA_SM", "--dry-run"]) == 3
    assert "不能再指定店铺" in capsys.readouterr().out
    assert env.run(["--region", "ALL", "--stores", "BOCINA_SM"],
                   {"MX": plan(("report", child(0)))}) == 3
    _, doc, _ = env.last()
    assert doc["exit"] == 3 and doc["regions"] == []
    assert not os.path.exists(run_monthly.LOCK)


# ------------------------------------------------------------ 一轮

def test_all_regions_run_in_one_log_with_region_on_every_line(env):
    code = env.run(["--region", "ALL", "--entrance", "scheduled"], {
        "MX": plan(("download", child(0, "# [1/2] BOCINA_SM\n  拿到了")),
                   ("report", child(0, "单店报表：BOCINA_SM"))),
        "CL": plan(("download", child(1, "# [1/2] Tecnohub_CL\n  [警告] 少一张")),
                   ("report", child(0, "单店报表：Tecnohub_CL"))),
    })
    assert code == 1                                  # 最坏的那个
    rid, doc, recs = env.last()
    assert len(runlog.list_runs(env.log_dir)) == 1, "一份日志"

    head = [r for r in recs if r["step"] == "启动" and not
            runlog.split_region(r["message"])[0]]
    text = "\n".join(r["message"] for r in head)
    assert "区域　　：全部国家，依次跑：墨西哥（MX） → 智利（CL）" in text
    assert "秘鲁（PE） config.json 里没有店铺，跳过" in text
    assert "命令 MX-1：" in text and "命令 CL-2：" in text

    # 段首醒目；段内每一行都带国家
    banners = [r["message"] for r in recs if "个国家：" in r["message"]]
    assert banners[0].startswith("[MX] ━━") and "第 1/2 个国家：墨西哥（MX）" in banners[0]
    assert banners[1].startswith("[CL] ━━")
    i_cl = next(i for i, r in enumerate(recs) if r["message"] == banners[1])
    i_mx = next(i for i, r in enumerate(recs) if r["message"] == banners[0])
    i_end = next(i for i, r in enumerate(recs) if r["message"] == "全部国家的结果：")
    assert all(r["message"].startswith("[MX] ") for r in recs[i_mx:i_cl])
    assert all(r["message"].startswith("[CL] ") for r in recs[i_cl:i_end])
    assert ("下载", "Tecnohub_CL", "警告", "[CL]   [警告] 少一张") in [
        (r["step"], r["store"], r["level"], r["message"]) for r in recs]

    # 最后的总结：每个国家一行，再加上跳过的
    tail = [r["message"] for r in recs[i_end:]]
    assert any(m.startswith("  墨西哥（MX）：退出码 0") for m in tail)
    assert any(m.startswith("  智利（CL）：退出码 1") for m in tail)
    assert "  秘鲁（PE）：没有店铺，跳过" in tail
    assert "本轮结束：退出码 1" in tail[-1]

    assert doc["region"] == "ALL" and doc["status"] == "done" and doc["exit"] == 1
    assert [(r["region"], r["exit"]) for r in doc["regions"]] == [("MX", 0), ("CL", 1)]
    assert doc["skipped_regions"] == ["PE"]
    assert [(s["region"], s["step"]) for s in doc["steps"]] == [
        ("MX", "download"), ("MX", "report"), ("CL", "download"), ("CL", "report")]
    assert {c["region"] for c in doc["commands"]} == {"MX", "CL"}

    # 整个月一份「！需要处理」（core/month_overview）：智利那条带「智利」，墨西哥没事就没有它的话
    assert env.alert("MX") == env.alert("CL") and os.path.exists(env.alert("CL"))
    with io.open(env.alert("CL"), encoding="utf-8") as f:
        text = f.read()
    assert "智利 —— 下载这一步有需要关注的项" in text
    assert not any("墨西哥 —— " in line for line in text.splitlines())
    assert doc["regions"][1]["alert"] == os.path.abspath(env.alert("CL"))
    assert not [n for n in os.listdir(env.out_dir) if n.startswith("！报表异常")]
    assert os.path.exists(os.path.join(env.out_dir, "本月总览_%s.xlsx" % MONTH))
    # 收尾干净
    assert not os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir) is None

    st = runlog.run_state(rid, env.log_dir)
    assert st["region"] == "ALL"
    assert "  全部国家 %s  " % MONTH in runlog.run_label(
        runlog.list_runs(env.log_dir)[0], env.log_dir)


def test_a_hard_failure_in_one_region_does_not_stop_the_next(env, tmp_path):
    ran = tmp_path / "cl_report_ran"
    code = env.run(["--region", "ALL"], {
        "MX": plan(("download", child(2, "紫鸟客户端错误：起不来")),
                   ("report", child(0))),
        "CL": plan(("download", child(0)),
                   ("report", [PY, "-c", "open(%r, 'w').close()" % str(ran)])),
    })
    assert code == 2
    assert ran.exists(), "墨西哥下载失败，智利照跑"
    _, doc, recs = env.last()
    assert [(r["region"], r["exit"]) for r in doc["regions"]] == [("MX", 2), ("CL", 0)]
    assert [s["step"] for s in doc["regions"][0]["steps"]] == ["download"]
    assert os.path.exists(env.alert("MX"))
    with io.open(env.alert("MX"), encoding="utf-8") as f:
        assert "墨西哥 —— 下载失败" in f.read()
    assert "墨西哥：下载失败" in recs[-1]["message"]
    assert all(u.startswith("[MX] ") for u in doc["urgent"])


def test_deadline_in_one_region_and_worst_code_ordering(env):
    assert run_monthly.worst_code([0, 1]) == 1
    assert run_monthly.worst_code([1, 7]) == 7
    assert run_monthly.worst_code([7, 2]) == 2       # 和单个国家里一样：7 让位给 2
    assert run_monthly.worst_code([4, 2, 0]) == 4
    assert run_monthly.worst_code([]) == 0
    assert run_monthly.worst_code([None]) == 2       # 认不出来的按硬失败算


def test_a_region_with_a_broken_batch_is_skipped_others_run(env, cfg):
    cfg["batches"]["CL"] = ["NOT_ALLOWED"]
    code = env.run(["--region", "ALL"], {"MX": plan(("report", child(0)))})
    assert code == 3
    _, doc, recs = env.last()
    assert [(r["region"], r["exit"]) for r in doc["regions"]] == [("MX", 0), ("CL", 3)]
    assert os.path.exists(env.alert("CL"))
    assert any(r["message"].startswith("[CL] 店铺不在白名单") for r in recs)


def test_config_unreadable_is_exit_4_without_lock(env, monkeypatch):
    def broken():
        raise store_config.StoreConfigError("坏了")
    monkeypatch.setattr(store_config, "_stores", broken)
    assert env.run(["--region", "ALL"]) == 4
    _, doc, _ = env.last()
    assert doc["exit"] == 4 and "坏了" in doc["urgent"][0]
    assert os.path.exists(env.alert("MX"))
    assert not os.path.exists(run_monthly.LOCK)


def test_no_store_anywhere_is_exit_4(env, cfg):
    cfg.update(allowed=[], batch=[], regions={}, batches={})
    assert env.run(["--region", "ALL"]) == 4


def test_the_lock_is_shared_with_single_region_runs(env):
    os.makedirs(env.log_dir)
    with io.open(run_monthly.LOCK, "w", encoding="utf-8") as f:
        f.write("pid=%d 起于 x" % os.getpid())
    assert env.run(["--region", "ALL"], {"MX": plan(("report", child(0))),
                                         "CL": plan(("report", child(0)))}) == 5


def test_stop_during_the_first_region_stops_everything(env, tmp_path):
    ran = tmp_path / "cl_ran"

    def stopper():
        for _ in range(200):
            cur = runlog.active_run(env.log_dir)
            if cur and cur.get("current_region") == "MX":
                time.sleep(0.6)
                runlog.request_stop(cur["run_id"], "测试", env.log_dir)
                return
            time.sleep(0.05)

    t = threading.Thread(target=stopper)
    t.start()
    code = env.run(["--region", "ALL"], {
        "MX": plan(("download", [PY, "-c", "import time; time.sleep(30)"])),
        "CL": plan(("report", [PY, "-c", "open(%r, 'w').close()" % str(ran)])),
    })
    t.join()
    assert code == 6
    assert not ran.exists()
    _, doc, recs = env.last()
    assert doc["status"] == "aborted"
    assert [(r["region"], r["status"]) for r in doc["regions"]] == [("MX", "aborted")]
    assert any("智利（CL）：没跑（前面被中止了）" in r["message"] for r in recs)
    assert not os.path.exists(env.alert("MX")) and not os.path.exists(env.alert("CL"))


def test_current_json_says_which_region_is_running(env, tmp_path):
    seen = {}

    def watcher():
        for _ in range(400):
            cur = runlog.active_run(env.log_dir)
            if cur and cur.get("current_region"):
                seen.setdefault(cur["current_region"], dict(cur))
            if len(seen) == 2:
                return
            time.sleep(0.02)

    t = threading.Thread(target=watcher)
    t.start()
    env.run(["--region", "ALL"], {
        "MX": plan(("report", [PY, "-c", "import time; time.sleep(1)"])),
        "CL": plan(("report", [PY, "-c", "import time; time.sleep(1)"]))})
    t.join()
    assert seen["MX"]["region"] == "ALL" and seen["MX"]["region_index"] == 1
    assert seen["CL"]["region_index"] == 2 and seen["CL"]["region_count"] == 2


# ------------------------------------------------------------ runlog

def test_tag_and_split_region():
    assert runlog.tag_region("CL", "a\nb") == "[CL] a\n[CL] b"
    assert runlog.split_region("[CL] 开始") == ("CL", "开始")
    assert runlog.split_region("[OK] 不是国家") == (None, "[OK] 不是国家")
    assert runlog.split_region("开始") == (None, "开始")
    rec = runlog.as_record("2026-10-13 14:05:01 | 定时 | 下载 | A | 警告 | [MX] x")
    assert (rec["region"], rec["body"]) == ("MX", "x")
    assert runlog.matches(rec, region="MX") and not runlog.matches(rec, region="CL")
    assert runlog.matches(rec, text="X") and not runlog.matches(rec, text="zzz")
    assert runlog.matches(rec, text="a")                 # 店名里也找


def _write_all_run(log_dir, run_id, regions):
    os.makedirs(runlog.runs_dir(log_dir), exist_ok=True)
    recs = []
    for region, rows in regions:
        batch = os.path.join(log_dir, "batch_%s_%s.json" % (run_id, region))
        with open(batch, "w", encoding="utf-8") as f:
            json.dump({"stores": rows}, f, ensure_ascii=False)
        recs.append({"region": region, "batch_summary": batch, "exit": 1,
                     "steps": []})
    runlog.write_json(runlog.run_paths(run_id, log_dir)["json"],
                      {"run_id": run_id, "month": MONTH, "stores": "all",
                       "region": "ALL", "regions": recs, "status": "done",
                       "exit": 1})
    open(runlog.run_paths(run_id, log_dir)["log"], "w").close()


def _login(store):
    return {"store": store, "status": "failed",
            "missing": [{"form": "ventas", "cause": "PAGE_BLOCKED"}]}


def test_human_actions_from_an_all_regions_run(tmp_path, cfg):
    d = str(tmp_path / "logs")
    _write_all_run(d, "20261014_020000_scheduled",
                   [("MX", [_login("BOCINA_SM")]), ("CL", [_login("Andes_CL")])])
    got = runlog.pending_human_actions(d)
    assert {(a["store"], a["region"]) for a in got["actions"]} == {
        ("BOCINA_SM", "MX"), ("Andes_CL", "CL")}
    # 之后单独补跑了智利那家（好了）：只剩墨西哥那条
    os.makedirs(runlog.runs_dir(d), exist_ok=True)
    batch = os.path.join(d, "batch_fix.json")
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": [{"store": "Andes_CL", "status": "ok"}]}, f)
    runlog.write_json(runlog.run_paths("20261014_090000_manual", d)["json"],
                      {"run_id": "20261014_090000_manual", "month": MONTH,
                       "region": "CL", "stores": ["Andes_CL"],
                       "batch_summary": batch, "status": "done", "exit": 0})
    open(runlog.run_paths("20261014_090000_manual", d)["log"], "w").close()
    got = runlog.pending_human_actions(d)
    assert [a["store"] for a in got["actions"]] == ["BOCINA_SM"]


def test_exit2_text_names_the_failing_region():
    doc = {"region": "ALL", "regions": [
        {"region": "MX", "exit": 0, "steps": []},
        {"region": "CL", "exit": 2, "steps": [{"step": "download", "code": 2}]}]}
    assert runlog.exit2_text(doc).startswith("智利：下载失败")


def test_live_status_says_country_step_store_and_time():
    now = datetime.datetime(2026, 10, 15, 3, 12, 5)
    st = {"state": "running", "run_id": "20261015_020000_scheduled",
          "month": MONTH, "summary": {"region": "ALL"},
          "current": {"region": "ALL", "current_region": "CL",
                      "region_index": 2, "region_count": 3}}
    recs = [runlog.as_record(x) for x in (
        "2026-10-15 03:00:00 | 定时 | 下载 | Tecnohub_CL | 信息 | [CL] # [2/2] Tecnohub_CL",
        "2026-10-15 03:10:00 | 定时 | 下载 | Tecnohub_CL | 信息 | [CL]   账单页")]
    text = runlog.live_status(st, recs, now=now)
    assert text == ("智利（CL），第 2/3 个国家  ·  2026-09  ·  下载  ·  "
                    "Tecnohub_CL（第 2/2 家）  ·  已运行 1:12:05")
    single = dict(st, summary={"region": "MX"}, current={})
    assert runlog.live_status(single, [], now=now).startswith("墨西哥（MX）  ·  2026-09")
    assert runlog.live_status({"state": "finished"}, recs) is None
    assert runlog.live_status(None) is None


# ------------------------------------------------------------ 勾选的几个（窗口按平台 / 国家勾）

def test_a_chosen_list_parses_in_the_all_order():
    assert run_monthly.parse_args(["--region", "cl,mx"]).region == "MX,CL"
    assert run_monthly.parse_args(["--region", "CL, MX,CL"]).region == "MX,CL"
    with pytest.raises(SystemExit):
        run_monthly.parse_args(["--region", "MX,XX"])
    with pytest.raises(SystemExit):
        run_monthly.parse_args(["--region", "MX,ALL"])


def test_a_chosen_list_with_a_store_subset_in_one_region(env, capsys):
    """墨西哥整个 + 智利只一家：墨西哥的命令和单独跑一样，智利那段只跑这一家、不出汇总。"""
    mx = _dry(env, capsys, [])
    lines = _dry(env, capsys, ["--region", "CL,MX", "--stores", "Andes_CL"])
    assert lines[:len(mx)] == mx
    cl = lines[len(mx):]
    assert cl and all("--stores Andes_CL" in l for l in cl)
    assert any("run_reports" in l and "--no-combined" in l for l in cl)
    assert env.run(["--region", "MX,CL", "--stores", "Andes_CL", "--dry-run"]) == 0
    assert "只跑 1 家：Andes_CL" in capsys.readouterr().out


def test_a_store_outside_the_chosen_regions_is_exit_3(env, capsys):
    assert env.run(["--region", "CL", "--stores", "BOCINA_SM", "--dry-run"]) == 3
    capsys.readouterr()
    assert env.run(["--region", "MX,PE", "--stores", "Andes_CL", "--dry-run"]) == 3
    assert "不在所选的区域里" in capsys.readouterr().out


def test_a_chosen_list_runs_like_all_and_marks_partial_regions(env):
    code = env.run(["--region", "MX,CL", "--stores", "Andes_CL"], {
        "MX": plan(("report", child(0, "单店报表：BOCINA_SM"))),
        "CL": plan(("report", child(0, "单店报表：Andes_CL")))})
    assert code == 0
    _, doc, recs = env.last()
    assert doc["region"] == "MX,CL" and doc["stores"] == ["Andes_CL"]
    assert [(r["region"], r["all_stores"], r["stores"]) for r in doc["regions"]] == [
        ("MX", True, MX), ("CL", False, ["Andes_CL"])]
    heads = [r["message"] for r in recs if "区域　　：" in (r["message"] or "")]
    assert heads and heads[0].startswith("区域　　：所选的 2 个，依次跑")
    assert runlog.region_name(doc["region"]) == "墨西哥、智利"
    assert runlog.is_multi(doc["region"]) and not runlog.is_multi("FA-CL")


def test_human_actions_of_a_partial_region_do_not_hide_older_ones(tmp_path, cfg):
    """勾选的几个里只跑了智利一家：智利更早那轮的待办（别的店）不能因此消失。"""
    d = str(tmp_path / "logs")
    _write_all_run(d, "20261014_020000_scheduled", [("CL", [_login("Tecnohub_CL")])])
    os.makedirs(runlog.runs_dir(d), exist_ok=True)
    batch = os.path.join(d, "batch_part.json")
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": [{"store": "Andes_CL", "status": "ok"}]}, f)
    rid = "20261014_090000_manual"
    runlog.write_json(runlog.run_paths(rid, d)["json"],
                      {"run_id": rid, "month": MONTH, "region": "MX,CL",
                       "stores": ["Andes_CL"], "status": "done", "exit": 0,
                       "regions": [{"region": "CL", "all_stores": False,
                                    "batch_summary": batch}]})
    open(runlog.run_paths(rid, d)["log"], "w").close()
    got = runlog.pending_human_actions(d)
    assert [a["store"] for a in got["actions"]] == ["Tecnohub_CL"]
