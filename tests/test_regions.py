# -*- coding: utf-8 -*-
"""多国家（区域）：墨西哥一字不变，智利永远带显式店铺清单。

config.json 一律用桩（store_config._stores），不读真文件。
"""
import io
import json
import os
import sys

import pytest

from core import region_choice
from platforms.mercadolibre import run_batch
from core import run_monthly
from platforms.mercadolibre import run_reports
from core import runlog
from core import store_config

PY = sys.executable
MONTH = "2026-09"
CL = ["Tecnohub_CL", "Andes_CL"]
MX = ["BOCINA_SM", "EWTTO_SM"]

CFG = {"default": "BOCINA_SM",
       "allowed": MX + CL + ["UNIT_PW01"],
       "batch": MX,
       "regions": {"Tecnohub_CL": "CL", "Andes_CL": "cl"},
       "batches": {"CL": CL}}


@pytest.fixture
def cfg(monkeypatch):
    """换掉 config.json 的 stores 段。返回那个 dict，测试里可以就地改。"""
    d = json.loads(json.dumps(CFG))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


# ------------------------------------------------------------ store_config

def test_region_of_defaults_to_mexico(cfg):
    assert store_config.region_of("Tecnohub_CL") == "CL"
    assert store_config.region_of("Andes_CL") == "CL"          # 小写也认
    assert store_config.region_of("BOCINA_SM") == "MX"
    assert store_config.region_of("没配过的店") == "MX"
    del cfg["regions"]
    assert store_config.region_of("Tecnohub_CL") == "MX"       # 只认 config


def test_stores_for_region_uses_the_configured_batches(cfg):
    assert store_config.stores_for_region("MX") == MX
    assert store_config.stores_for_region("mx") == MX
    assert store_config.stores_for_region("CL") == CL


def test_stores_for_region_falls_back_to_allowed_by_region(cfg):
    cfg["batch"] = []
    del cfg["batches"]
    # 墨西哥：白名单里没标成别国的店（不是 batch_stores()，那会把智利店也带上）
    assert store_config.stores_for_region("MX") == MX + ["UNIT_PW01"]
    assert store_config.stores_for_region("CL") == CL
    assert store_config.batch_stores() == CFG["allowed"]       # 旧函数不变


def test_a_foreign_store_in_a_batch_is_a_config_error(cfg):
    cfg["batch"] = MX + ["Tecnohub_CL"]
    with pytest.raises(store_config.StoreConfigError) as e:
        store_config.stores_for_region("MX")
    assert "Tecnohub_CL" in str(e.value)
    cfg["batches"]["CL"] = ["NOT_ALLOWED"]
    with pytest.raises(store_config.StoreNotAllowed):
        store_config.stores_for_region("CL")


def test_regions_lists_mexico_first(cfg):
    assert store_config.regions() == ["MX", "CL"]
    cfg["regions"] = {}
    cfg["batches"] = {}
    assert store_config.regions() == ["MX"]
    cfg["allowed"] = list(CL)
    cfg["regions"] = {"Tecnohub_CL": "CL", "Andes_CL": "CL"}
    assert store_config.regions() == ["CL"]


def test_group_by_region_keeps_order(cfg):
    assert store_config.group_by_region(["Tecnohub_CL", "BOCINA_SM", "Andes_CL"]) \
        == [("CL", ["Tecnohub_CL", "Andes_CL"]), ("MX", ["BOCINA_SM"])]


# ------------------------------------------------------------ run_monthly：命令

def R(name):
    return os.path.join(run_monthly.ROOT, name)


def test_mexico_child_commands_are_byte_for_byte_unchanged():
    """区域出现之前的命令，原样写死在这里。墨西哥一个字都不能变。"""
    dl = ["py", "-u", "-m", "platforms.mercadolibre.run_batch", "--month", MONTH]
    rp = ["py", "-u", "-m", "platforms.mercadolibre.run_reports", "--month", MONTH]
    cases = [
        (("both", None, False), [dl, rp + ["--out-dir", "OUT"]]),
        (("download", None, False), [dl]),
        (("report", None, False), [rp + ["--out-dir", "OUT"]]),
        (("collect", None, False),
         [["py", "-u", "-m", "platforms.mercadolibre.run_batch", "--collect-only", "--month", MONTH]]),
        (("both", ["A", "B"], False),
         [dl + ["--stores", "A", "B"],
          rp + ["--stores", "A", "B", "--no-combined", "--out-dir", "OUT"]]),
        (("both", ["A"], True), [dl + ["--stores", "A"], rp + ["--out-dir", "OUT"]]),
        (("collect", ["A"], False),
         [["py", "-u", "-m", "platforms.mercadolibre.run_batch", "--collect-only", "--month", MONTH,
           "--stores", "A"]]),
    ]
    for (step, stores, all_), want in cases:
        for kw in ({}, {"region": "MX"}, {"region": "MX", "all_stores": MX}):
            got = [s["cmd"] for s in run_monthly.build_plan(
                step, MONTH, stores, "OUT", py="py", reports_all=all_, **kw)]
            assert got == want, (step, stores, all_, kw)


def test_chile_always_passes_an_explicit_store_list():
    p = run_monthly.build_plan("both", MONTH, None, "OUT", py="py",
                               region="CL", all_stores=CL)
    dl, rp = p[0]["cmd"], p[1]["cmd"]
    assert dl == ["py", "-u", "-m", "platforms.mercadolibre.run_batch", "--month", MONTH,
                  "--stores"] + CL
    # 全部智利店：run_reports 自己按区域取店，并且要出（智利的）全店汇总
    assert rp == ["py", "-u", "-m", "platforms.mercadolibre.run_reports", "--month", MONTH,
                  "--region", "CL", "--out-dir", "OUT"]

    sub = run_monthly.build_plan("both", MONTH, ["Andes_CL"], "OUT", py="py",
                                 region="CL", all_stores=CL)
    assert sub[0]["cmd"][-2:] == ["--stores", "Andes_CL"]
    assert sub[1]["cmd"][4:] == ["--month", MONTH, "--region", "CL", "--stores",
                                 "Andes_CL", "--no-combined", "--out-dir", "OUT"]

    co = run_monthly.build_plan("collect", MONTH, None, "OUT", py="py",
                                region="CL", all_stores=CL)
    assert co[0]["cmd"][-3:] == ["--stores"] + CL

    again = run_monthly.build_plan("both", MONTH, ["Andes_CL"], "OUT", py="py",
                                   region="CL", all_stores=CL, reports_all=True)
    assert "--stores" not in again[1]["cmd"] and "--region" in again[1]["cmd"]

    with pytest.raises(ValueError):            # 绝不退回"不带店铺"
        run_monthly.build_plan("download", MONTH, None, "OUT", region="CL",
                               all_stores=[])


def test_region_flag_parses_and_defaults_to_mexico():
    assert run_monthly.parse_args([]).region == "MX"
    assert run_monthly.parse_args(["--region", "cl"]).region == "CL"
    with pytest.raises(SystemExit):
        run_monthly.parse_args(["--region", "AR"])


def test_alert_file_name_per_region():
    assert run_monthly.alert_name("2026-09") == "！报表异常_202609.txt"
    assert run_monthly.alert_name("2026-09", "MX") == "！报表异常_202609.txt"
    assert run_monthly.alert_name("2026-09", "CL") == "！报表异常_CL_202609.txt"


# ------------------------------------------------------------ run_monthly：一轮

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

        def run(self, argv, plan=None):
            args = run_monthly.parse_args(
                ["--month", MONTH, "--out-dir", self.out_dir] + argv)
            return run_monthly.execute(args, plan=plan)

        def last(self):
            rid = runlog.list_runs(self.log_dir)[0]["run_id"]
            p = runlog.run_paths(rid, self.log_dir)
            with io.open(p["log"], encoding="utf-8") as f:
                recs = [runlog.parse_line(x) for x in f.read().splitlines()]
            return rid, runlog.read_json(p["json"]), recs

        def alert(self, region="MX"):
            return os.path.join(self.out_dir, run_monthly.alert_name(MONTH, region))

    return Env()


def child(code, text=""):
    return [PY, "-c", "import sys\nprint(%r, flush=True)\nsys.exit(%d)\n"
            % (text, code)]


def test_chile_dry_run_prints_the_explicit_store_list(env, capsys):
    assert env.run(["--region", "CL", "--dry-run"]) == 0
    out = capsys.readouterr().out
    assert "区域　　：智利（CL）" in out
    assert "-m platforms.mercadolibre.run_batch --month %s --stores Tecnohub_CL Andes_CL" % MONTH in out
    assert "-m platforms.mercadolibre.run_reports --month %s --region CL --out-dir" % MONTH in out


def test_chile_run_records_region_in_log_summary_and_alert(env):
    code = env.run(["--region", "CL", "--entrance", "scheduled"], [
        {"key": "download", "label": "d", "cmd": child(2, "boom")},
        {"key": "report", "label": "r", "cmd": child(0)}])
    assert code == 2
    rid, doc, recs = env.last()
    head = "\n".join(r["message"] for r in recs if r["step"] == "启动")
    assert "区域　　：智利（CL）" in head
    assert "智利（CL）的店铺" in head
    assert doc["region"] == "CL" and doc["stores"] == "all"
    assert os.path.exists(env.alert("CL"))
    assert not os.path.exists(env.alert("MX"))      # 墨西哥那份不能被它占掉
    with io.open(env.alert("CL"), encoding="utf-8") as f:
        assert f.read().startswith("智利（CL） %s 月报生成异常" % MONTH)
    st = runlog.run_state(rid, env.log_dir)
    assert st["region"] == "CL"
    assert "  智利 %s  " % MONTH in runlog.run_label(
        runlog.list_runs(env.log_dir)[0], env.log_dir)


def test_mexico_run_keeps_the_old_alert_name_and_records_region(env):
    assert env.run(["--step", "download"], [
        {"key": "download", "label": "d", "cmd": child(2)}]) == 2
    _, doc, recs = env.last()
    assert doc["region"] == "MX"
    assert os.path.exists(os.path.join(env.out_dir, "！报表异常_202609.txt"))
    assert any(r["message"] == "区域　　：墨西哥（MX）" for r in recs)


@pytest.mark.parametrize("region,stores", [
    ("CL", ["Tecnohub_CL", "BOCINA_SM"]),
    ("MX", ["BOCINA_SM", "Andes_CL"]),
])
def test_stores_from_another_region_exit_3_before_anything_runs(
        env, region, stores, capsys):
    argv = ["--region", region, "--stores"] + stores
    assert env.run(argv + ["--dry-run"]) == 3
    assert "不属于" in capsys.readouterr().out

    code = env.run(argv, [{"key": "download", "label": "d", "cmd": child(0)}])
    assert code == 3
    _, doc, recs = env.last()
    assert doc["status"] == "failed" and doc["exit"] == 3 and doc["steps"] == []
    bad = stores[1]
    assert any(bad in r["message"] and "不属于" in r["message"] for r in recs)
    assert not os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir) is None
    assert os.path.exists(env.alert(region))


def test_chile_without_any_chilean_store_is_a_config_error(env, cfg):
    cfg["regions"] = {}
    cfg["batches"] = {}
    assert env.run(["--region", "CL"]) == 4
    _, doc, recs = env.last()
    assert doc["exit"] == 4 and doc["steps"] == []


def test_lock_is_shared_by_both_regions(env):
    os.makedirs(env.log_dir)
    with io.open(run_monthly.LOCK, "w", encoding="utf-8") as f:
        f.write("pid=%d 起于 x" % os.getpid())          # 一轮墨西哥正在跑
    assert env.run(["--region", "CL"], [
        {"key": "report", "label": "r", "cmd": child(0)}]) == 5


# ------------------------------------------------------------ run_reports

def test_combined_workbook_name_per_region():
    assert run_reports.combined_name("2026-09") == "MX_ML_全店汇总_2026-09.xlsx"
    assert run_reports.combined_name("2026-09", "MX") == "MX_ML_全店汇总_2026-09.xlsx"
    assert run_reports.combined_name("2026-09", "CL") == "CL_ML_全店汇总_2026-09.xlsx"


@pytest.mark.parametrize("region,stores,combined", [
    ("MX", MX, "MX_ML_全店汇总_2026-08.xlsx"),
    ("CL", CL, "CL_ML_全店汇总_2026-08.xlsx"),
])
def test_run_reports_picks_the_region_and_names_the_files(
        monkeypatch, tmp_path, cfg, region, stores, combined):
    outs = []

    def build(pairs, month, out, **kw):
        outs.append((len(pairs), os.path.basename(out)))
        return {"ok": True, "skipped": None, "returncode": 0, "seconds": 1,
                "wrote": True, "stale": False, "checks": [], "checks_failed": False,
                "data_incomplete": False, "formula_errors": False, "output": "",
                "out": out}
    monkeypatch.setattr(run_reports, "resolve",
                        lambda names, month=None: ([(n, "/d/" + n) for n in names], []))
    monkeypatch.setattr(run_reports.financial_report, "has_libreoffice", lambda: True)
    monkeypatch.setattr(run_reports.financial_report, "build", build)
    argv = ["run_reports.py", "--month", "2026-08", "--out-dir", str(tmp_path)]
    if region != "MX":
        argv += ["--region", region]
    monkeypatch.setattr(sys, "argv", argv)
    assert run_reports.main() == 0
    # 单店报表的文件名不变；汇总表按区域起名
    assert outs == [(1, "%s_2026-08.xlsx" % s) for s in stores] + [
        (len(stores), combined)]


def test_run_reports_explicit_stores_ignore_region_lookup(cfg):
    assert run_reports.store_names(["Andes_CL"], "CL") == ["Andes_CL"]
    assert run_reports.store_names(None, "CL") == CL
    assert run_reports.store_names(None) == MX


# ------------------------------------------------------------ run_batch

def test_cleaning_applies_only_with_a_mexican_store(cfg):
    assert run_batch.cleaning_applies(MX)
    assert run_batch.cleaning_applies(CL + ["BOCINA_SM"])
    assert not run_batch.cleaning_applies(CL)
    assert run_batch.cleaning_state(run_batch.region_skip()) == "skipped"
    assert run_batch.batch_verdict(
        [{"store": "A", "status": "ok"}], run_batch.region_skip())["exit_code"] == 0


def test_all_chilean_batch_skips_cleaning_and_says_so(
        monkeypatch, tmp_path, cfg, capsys):
    logs = str(tmp_path / "logs")
    monkeypatch.setattr(run_batch, "LOG_DIR", logs)
    monkeypatch.setattr(run_batch.time, "sleep", lambda s: None)
    monkeypatch.setattr(run_batch.store_config, "describe", lambda: "test")
    monkeypatch.setattr(run_batch.pending_store, "prune", lambda: [])
    monkeypatch.setattr(run_batch.pending_store, "load", lambda: [])

    def no_cleaning(**kw):
        raise AssertionError("智利批次不该清洗")
    monkeypatch.setattr(run_batch.transform, "build", no_cleaning)

    class Z(object):
        def start(self, restart=True):
            pass
    monkeypatch.setattr(run_batch, "ZiniaoClient", Z)

    def fake_run_store(client, name, args, out_dir=None, close_when_done=True,
                       defer_slow=False):
        return {"store": name, "out_dir": out_dir, "billing": {}, "sales": [],
                "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
                "expected": [], "missing": [], "flags": {}}
    monkeypatch.setattr(run_batch.run_downloads, "run_store", fake_run_store)
    monkeypatch.setattr(run_batch.run_downloads, "print_summary", lambda r: [])
    monkeypatch.setattr(run_batch, "finalize_diagnosis", lambda *a, **k: None)
    monkeypatch.setattr("sys.argv", ["run_batch.py", "--pace", "0",
                                     "--no-late-collect", "--out",
                                     str(tmp_path / "dl"), "--stores"] + CL)

    assert run_batch.main() == 0
    out = capsys.readouterr().out
    assert "清洗目前只支持墨西哥，其他国家的店铺跳过这一步；不影响报表" in out
    path = out.rstrip().splitlines()[-1].split(" ", 1)[1]
    doc = json.load(io.open(path, encoding="utf-8"))
    assert doc["transform"]["skipped"] is True
    assert doc["transform"]["reason"] == "region"
    assert doc["cleaning"] == "skipped" and doc["exit_code"] == 0
    # 运行器读这份汇总：不当成清洗失败，也不告警
    urgent, attention = run_monthly.describe_batch(path)
    assert urgent == [] and attention == []


# ------------------------------------------------------------ runlog

def test_runner_argv_region_is_optional():
    assert runlog.runner_argv("scheduled") == ["--entrance", "scheduled"]
    assert runlog.runner_argv("scheduled", region="CL") == [
        "--entrance", "scheduled", "--region", "CL"]
    a = run_monthly.parse_args(runlog.runner_argv("manual", region="CL"))
    assert a.region == "CL"


def _write_run(log_dir, run_id, stores, rows, region=None):
    os.makedirs(runlog.runs_dir(log_dir), exist_ok=True)
    batch = os.path.join(log_dir, "batch_%s.json" % run_id)
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": rows}, f, ensure_ascii=False)
    doc = {"run_id": run_id, "month": MONTH, "stores": stores,
           "status": "done", "exit": 1, "batch_summary": batch}
    if region:
        doc["region"] = region
    runlog.write_json(runlog.run_paths(run_id, log_dir)["json"], doc)
    open(runlog.run_paths(run_id, log_dir)["log"], "w").close()


def _login(store):
    return {"store": store, "status": "failed",
            "missing": [{"form": "ventas", "cause": "PAGE_BLOCKED"}]}


def test_a_full_chile_run_does_not_hide_mexican_actions(tmp_path):
    d = str(tmp_path / "logs")
    _write_run(d, "20261014_020000_scheduled", "all", [_login("BOCINA_SM")])
    _write_run(d, "20261014_090000_manual", "all", [_login("Andes_CL")],
               region="CL")
    got = runlog.pending_human_actions(d)
    assert {(a["store"], a["region"]) for a in got["actions"]} == {
        ("BOCINA_SM", "MX"), ("Andes_CL", "CL")}
    # 墨西哥再跑一整轮（好了）：只有更早那轮墨西哥的过时，智利的还在
    _write_run(d, "20261014_120000_manual", "all",
               [{"store": "BOCINA_SM", "status": "ok", "missing": []}])
    got = runlog.pending_human_actions(d)
    assert [a["store"] for a in got["actions"]] == ["Andes_CL"]


# ------------------------------------------------------------ region_choice

def test_task_names_and_arguments():
    assert region_choice.task_name("MX") == "MX月度报表"     # 生产上的名字
    assert region_choice.task_name("CL") == "CL月度报表"
    assert region_choice.rehearsal_name("CL") == "CL月度报表-试跑"
    assert region_choice.task_args("MX") == "--entrance scheduled --region MX"
    assert region_choice.task_args("CL") == "--entrance scheduled --region CL"
    assert region_choice.rehearsal_args("CL") == \
        "--entrance scheduled --region CL --step report"
    for r in ("MX", "CL"):
        a = run_monthly.parse_args(region_choice.task_args(r).split())
        assert (a.entrance, a.region, a.month, a.stores) == ("scheduled", r, None, None)
    assert region_choice.args_region("--entrance scheduled") == "MX"   # 旧任务
    assert region_choice.args_region("--entrance scheduled --region CL") == "CL"


def test_schedule_clash():
    assert region_choice.clash((15, "02:00"), (15, "02:30"))
    assert region_choice.clash((15, "02:00"), (15, "05:00"))   # 3 小时也撞
    assert region_choice.clash((15, "02:00"), (15, "08:00")) is None
    assert region_choice.clash((15, "02:00"), (16, "02:00")) is None
    assert region_choice.clash((15, "02:00"), None) is None
    st = {"DAY": "15", "TIME": "02:10"}
    assert "15 号" in region_choice.clash((15, "02:00"), st)
    assert region_choice.clash((15, "02:00"), {"DAY": "?", "TIME": "?"}) is None


def test_settings_and_labels():
    assert region_choice.remembered({}) == "MX"
    assert region_choice.remembered({"region": "cl"}) == "CL"
    assert region_choice.remembered({"region": "AR"}) == "MX"
    assert region_choice.from_label("智利") == "CL"
    assert region_choice.from_label("墨西哥") == "MX"
    assert region_choice.labels() == ["墨西哥", "智利", "秘鲁"]
    assert region_choice.window_title("MX") == "meli-mx-pipeline —— 墨西哥店铺报表"
    assert region_choice.sched_key("sched_day", "MX") == "sched_day"
    assert region_choice.sched_key("sched_day", "CL") == "sched_day_CL"


def test_gui_store_list_and_preview_follow_the_region(cfg):
    assert region_choice.stores("CL") == CL
    assert region_choice.stores("MX") == MX
    assert region_choice.preview_cmd("py", "rr.py", MONTH, None, "MX") == [
        "py", "rr.py", "--month", MONTH, "--dry-run"]
    assert region_choice.preview_cmd("py", "rr.py", MONTH, ["Andes_CL"], "CL") == [
        "py", "rr.py", "--month", MONTH, "--dry-run", "--region", "CL",
        "--stores", "Andes_CL"]


def test_rerun_plan_derives_region_from_the_store(cfg):
    a = lambda s, rerun=True: {"store": s, "title": "要登录", "rerun": rerun}
    p = region_choice.rerun_plan([a("Andes_CL"), a("Tecnohub_CL"), a("Andes_CL")])
    assert (p["region"], p["stores"]) == ("CL", ["Andes_CL", "Tecnohub_CL"])
    p = region_choice.rerun_plan([a("BOCINA_SM"), a("EWTTO_SM", rerun=False)])
    assert (p["region"], p["stores"], p["useless"]) == (
        "MX", ["BOCINA_SM"], ["EWTTO_SM（要登录）"])
    # stores of two countries: one run over both, in the ALL order
    p = region_choice.rerun_plan([a("Andes_CL"), a("BOCINA_SM")])
    assert p["region"] == "MX,CL" and len(p["groups"]) == 2
    assert region_choice.action_line(a("Andes_CL")) == "Andes_CL（美客多 智利） —— 要登录"
    assert region_choice.action_line(a("BOCINA_SM")) == "BOCINA_SM（美客多 墨西哥） —— 要登录"
