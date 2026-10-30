# -*- coding: utf-8 -*-
"""Automatic period reports after a monthly run (run_monthly.period_hook): the
quarter after its last month's report step, the year in December — and never a
different exit code for the monthly run."""
import io
import os
import sys

import pytest

import period_data as P
from core import period_report as PR
from core import run_monthly
from core import runlog

PY = sys.executable


@pytest.fixture
def env(tmp_path, monkeypatch):
    root = P.configure(tmp_path, monkeypatch, v2=True)
    logs = tmp_path / "logs"
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(logs))
    monkeypatch.setattr(run_monthly, "LOCK", str(logs / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    monkeypatch.setattr(run_monthly, "known_stores", lambda: [])
    monkeypatch.setattr(run_monthly, "POLL", 0.1)

    class Env(object):
        log_dir = str(logs)
        out_dir = root

        def run(self, month, argv, plan):
            args = run_monthly.parse_args(["--month", month] + argv)
            return run_monthly.execute(args, plan=plan)

        def log_text(self):
            rid = runlog.list_runs(self.log_dir)[0]["run_id"]
            with io.open(runlog.run_paths(rid, self.log_dir)["log"], encoding="utf-8") as f:
                text = f.read()
            assert all(runlog.parse_line(x) for x in text.splitlines()), "every line in the standard format"
            return text

        def alerts(self, month):
            d = os.path.join(root, "monthly", month)
            return sorted(os.listdir(d)) if os.path.isdir(d) else []

    return Env()


def child(code):
    return [PY, "-c", "import sys\nsys.exit(%d)\n" % code]


def plan(*steps):
    return [{"key": k, "label": k, "cmd": c} for k, c in steps]


def quarter(store="S1", months=("2026-07", "2026-08", "2026-09")):
    for m in months:
        P.put(store, m, "20261001_000000")


Q3_MX = ("quarterly", "2026-Q3", "ML-MX", "MX_ML_2026-Q3_账单基准.xlsx")


def path(env, kind, period, market, name):
    return os.path.join(env.out_dir, kind, period, market, name)


@pytest.mark.parametrize("code", [0, 1, 2])
def test_quarter_end_generates_the_quarter_and_keeps_the_exit_code(env, code):
    quarter()
    assert env.run("2026-09", [], plan(("report", child(code)))) == code
    assert os.path.exists(path(env, *Q3_MX))
    text = env.log_text()
    assert "期间报表 美客多 墨西哥 · 2026-Q3（账单基准）：完整" in text
    assert not [a for a in env.alerts("2026-09") if a.startswith("！期间报表异常")]


def _alert_body(env, month="2026-09"):
    with io.open(os.path.join(env.out_dir, "monthly", month, "！需要处理_%s.txt" % month.replace("-", "")),
                 encoding="utf-8") as f:
        return f.read()


def test_incomplete_quarter_is_listed_in_the_month_alert_until_complete(env):
    """An incomplete period report is one item of the month's single 「！需要处理」
    (core/month_overview) — no ！期间报表异常_* file any more."""
    quarter(months=("2026-07", "2026-09"))
    assert env.run("2026-09", [], plan(("report", child(0)))) == 0
    assert not [a for a in env.alerts("2026-09") if a.startswith("！期间报表异常")]
    body = _alert_body(env)
    assert "美客多 墨西哥 · 2026-Q3（账单基准）：已生成，但有店铺月份没有完整版本" in body
    assert "S1 2026-08 无完整版本" in body and "来源：期间报表" in body
    assert "期间报表 美客多 墨西哥 · 2026-Q3（账单基准）：已生成，但有店铺月份没有完整版本" in env.log_text()
    P.put("S1", "2026-08", "20261012_000000")
    assert env.run("2026-09", ["--step", "report"], plan(("report", child(0)))) == 0
    # the other configured stores still have no report, so the file stays — but
    # the period item is gone
    assert "无完整版本" not in _alert_body(env)
    assert "期间报表 美客多 墨西哥 · 2026-Q3（账单基准）：完整" in env.log_text()


def test_a_crash_while_generating_never_changes_the_run(env, monkeypatch):
    quarter()

    def boom(*a, **k):
        raise RuntimeError("disk on fire")
    monkeypatch.setattr(PR, "generate", boom)
    assert env.run("2026-09", [], plan(("report", child(1)))) == 1
    text = env.log_text()
    assert "期间报表没生成（不影响这一轮的月报）" in text and "disk on fire" in text
    assert not [a for a in env.alerts("2026-09") if a.startswith("！期间报表异常")]
    body = _alert_body(env)
    assert "自动生成期间报表时出错" in body and "disk on fire" in body


def test_an_error_result_alerts_without_changing_the_code(env, monkeypatch):
    quarter()
    from core.period_report import workbook as WB
    monkeypatch.setattr(WB, "build", lambda *a, **k: 1 / 0)
    assert env.run("2026-09", [], plan(("report", child(0)))) == 0
    assert not [a for a in env.alerts("2026-09") if a.startswith("！期间报表异常")]
    assert "出错，没生成" in env.log_text() and "出错，没生成" in _alert_body(env)


def test_only_quarter_ends_and_december(env):
    quarter(months=("2026-10", "2026-11"))
    assert env.run("2026-10", [], plan(("report", child(0)))) == 0
    assert env.run("2026-11", [], plan(("report", child(0)))) == 0
    assert not os.path.exists(os.path.join(env.out_dir, "quarterly"))
    P.put("S1", "2026-12", "20270105_000000")
    assert env.run("2026-12", [], plan(("report", child(0)))) == 0
    assert os.path.exists(path(env, "quarterly", "2026-Q4", "ML-MX", "MX_ML_2026-Q4_账单基准.xlsx"))
    assert os.path.exists(path(env, "annual", "2026", "ML-MX", "MX_ML_2026_账单基准.xlsx"))


def test_no_report_step_no_period_report(env):
    quarter()
    assert env.run("2026-09", ["--step", "download"], plan(("download", child(0)))) == 0
    assert not os.path.exists(path(env, *Q3_MX))
    assert run_monthly.execute(run_monthly.parse_args(["--month", "2026-09", "--dry-run"])) == 0
    assert not os.path.exists(path(env, *Q3_MX))


def test_no_results_at_all_is_only_a_log_line(env):
    assert env.run("2026-09", [], plan(("report", child(0)))) == 0
    assert "这个期间还没有月度结果文件" in env.log_text()
    assert env.alerts("2026-09") == [] or not any(a.startswith("！期间") for a in env.alerts("2026-09"))


def test_several_units_each_get_their_own_and_one_problem_does_not_stop_the_next(env, monkeypatch):
    quarter()
    quarter("F1")
    real = PR.generate

    def flaky(period, market, *a, **k):
        if market == "ML-MX":
            raise RuntimeError("only Mexico breaks")
        return real(period, market, *a, **k)
    monkeypatch.setattr(PR, "generate", flaky)
    code = env.run("2026-09", ["--region", "MX,FA-CL"],
                   {"MX": plan(("report", child(0))), "FA-CL": plan(("report", child(1)))})
    assert code == 1
    assert os.path.exists(path(env, "quarterly", "2026-Q3", "FA-CL", "CL_FA_2026-Q3.xlsx"))
    text = env.log_text()
    assert "[MX] 期间报表没生成" in text and "[FA-CL] 期间报表 Falabella 智利 · 2026-Q3：完整" in text
