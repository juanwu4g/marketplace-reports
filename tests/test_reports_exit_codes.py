# -*- coding: utf-8 -*-
"""run_reports 的退出码约定，和 financial_report 对"写没写出来"的判断。

    0  全部写出且干净      1  全部写出、有要看的      2  有报表没写出来 / 没有店可用

build 一律用桩：这里测的是"拿到这样的结果该怎么判"，不是报表怎么算。
"""
import os
import sys

import pytest

from platforms.mercadolibre import financial_report
from platforms.mercadolibre import run_reports


def check(key="bridge_b", ok=False, severity="error", store="A", name="桥B"):
    return {"key": key, "store": store, "name": name, "ok": ok,
            "severity": severity, "skipped": False, "detail": "说明文字-%s" % key}


def result(**over):
    """一个"写出来了、全部通过"的 build 结果，按需改掉几项。"""
    r = {"ok": True, "skipped": None, "returncode": 0, "seconds": 1,
         "wrote": True, "stale": False, "checks": [], "checks_failed": False,
         "data_incomplete": False, "formula_errors": False, "output": ""}
    r.update(over)
    return r


WRITTEN_FAILED = dict(ok=False, returncode=1, checks_failed=True, checks=[check()])
STRICT_ABORTED = dict(ok=False, returncode=financial_report.RC_STRICT_ABORT,
                      wrote=False, checks_failed=True, checks=[check()])
CRASHED = dict(ok=False, returncode=financial_report.RC_CRASH, wrote=False)


# ---------------------------------------------------------------- classify

@pytest.mark.parametrize("fields,expected", [
    (dict(), "ok"),
    (WRITTEN_FAILED, "checks_failed"),
    (dict(ok=False, returncode=1, formula_errors=True), "formula_errors"),
    (dict(data_incomplete=True), "data_incomplete"),
    (STRICT_ABORTED, "strict_aborted"),
    (CRASHED, "crashed"),
    # 写完 xlsx 之后才崩：文件是新的，但这次生成没走完。
    (dict(ok=False, returncode=financial_report.RC_CRASH, wrote=True), "crashed"),
    (dict(ok=False, returncode=financial_report.RC_NO_STORE, wrote=False), "no_usable_store"),
    (dict(ok=False, returncode=None, wrote=False, timed_out=True), "timeout"),
    (dict(ok=False, returncode=1, wrote=False, stale=True), "stale"),
    (dict(ok=False, returncode=financial_report.RC_CRASH, wrote=False, stale=True,
          output="PermissionError: [Errno 13]"), "stale"),
    # 解释器默认的退出码 1 加上没有文件：不是"有校验未过"。
    (dict(ok=False, returncode=1, wrote=False), "not_written"),
    (dict(ok=False, skipped="找不到 report_build.py", wrote=False, returncode=None), "skipped"),
])
def test_classify(fields, expected):
    r = result(**fields)
    assert financial_report.classify(r) == expected
    assert financial_report.written(r) == (expected in financial_report.WRITTEN_OUTCOMES)
    assert expected in financial_report.OUTCOME_LABELS


# ------------------------------------------------------- run_reports.main

@pytest.fixture
def reports(monkeypatch, tmp_path):
    """跑 run_reports.main()：店铺 A、B 都有可用目录，build 按 label 给结果。"""
    state = {"by_store": {}, "unusable": [], "calls": []}

    def resolve(names, month=None):
        bad = dict(state["unusable"])
        return ([(n, str(tmp_path / n)) for n in names if n not in bad],
                [(n, bad[n]) for n in names if n in bad])

    def build(pairs, month, out, **kwargs):
        label = pairs[0][0] if len(pairs) == 1 else "ALL"
        state["calls"].append((label, kwargs.get("strict")))
        spec = state["by_store"].get(label, {})
        if isinstance(spec, Exception):
            raise spec
        return result(out=out, **spec)

    monkeypatch.setattr(run_reports.store_config, "stores_for_region",
                        lambda region="MX": ["A", "B"])
    monkeypatch.setattr(run_reports, "resolve", resolve)
    monkeypatch.setattr(run_reports.financial_report, "has_libreoffice", lambda: True)
    monkeypatch.setattr(run_reports.financial_report, "build", build)

    def run(*argv):
        monkeypatch.setattr(sys, "argv", ["run_reports.py", "--month", "2026-08",
                                          "--out-dir", str(tmp_path / "out")] + list(argv))
        return run_reports.main()

    state["run"] = run
    return state


def test_all_written_and_clean_is_zero(reports, capsys):
    assert reports["run"]() == 0
    assert [c[0] for c in reports["calls"]] == ["A", "B", "ALL"]
    assert "全部 3 份报表生成成功" in capsys.readouterr().out


def test_written_with_failed_checks_is_one(reports, capsys):
    reports["by_store"] = {"A": WRITTEN_FAILED, "ALL": WRITTEN_FAILED}
    assert reports["run"]() == 1
    out = capsys.readouterr().out
    assert "有校验未过" in out
    assert "A_2026-08.xlsx" in out          # 文件确实写出了，汇总里给文件名
    # 汇总里每份报表的未过校验只有一行压缩的，不重打说明文字。
    assert "说明文字-bridge_b" not in out
    assert out.count("桥B[bridge_b]（A）") == 2


def test_one_report_not_written_is_two(reports, capsys):
    reports["by_store"] = {"B": CRASHED}
    assert reports["run"]() == 2
    out = capsys.readouterr().out
    assert "生成出错" in out
    assert "1/3 份报表没有写出来：B" in out


def test_not_written_wins_over_failed_checks(reports):
    reports["by_store"] = {"A": WRITTEN_FAILED, "B": CRASHED}
    assert reports["run"]() == 2


def test_build_that_raises_is_two_and_the_rest_still_run(reports, capsys):
    reports["by_store"] = {"A": RuntimeError("磁盘满了")}
    assert reports["run"]() == 2
    assert [c[0] for c in reports["calls"]] == ["A", "B", "ALL"]
    out = capsys.readouterr().out
    assert "磁盘满了" in out and "生成出错" in out


def test_no_usable_store_is_two(reports, capsys):
    reports["unusable"] = [("A", "没有下载目录"), ("B", "没有下载目录")]
    assert reports["run"]() == 2
    assert reports["calls"] == []
    assert "没有一家店有可用的下载目录" in capsys.readouterr().out


def test_store_without_folder_means_its_report_was_not_written(reports):
    reports["unusable"] = [("B", "没有下载目录")]
    assert reports["run"]() == 2
    # 只出汇总时文件写出来了，只是少一家店：要看，但不是"没写出来"。
    assert reports["run"]("--no-per-store") == 1


def test_strict_abort_is_not_shown_as_a_written_file(reports, capsys):
    reports["by_store"] = {"A": STRICT_ABORTED}
    assert reports["run"]("--strict", "--no-combined") == 2
    assert reports["calls"] == [("A", True), ("B", True)]
    out = capsys.readouterr().out
    row = [l for l in out.splitlines() if l.startswith("  A ")][0]
    assert "严格模式未写出" in row and "有校验未过" not in row
    assert "A_2026-08.xlsx" not in row and row.rstrip().endswith("—")


def test_incomplete_ventas_is_one_and_labelled(reports, capsys):
    warn = check("ventas_incomplete", severity="warn", name="Ventas 导出数据不完整")
    reports["by_store"] = {"A": dict(data_incomplete=True, checks=[warn])}
    assert reports["run"]("--no-combined") == 1
    out = capsys.readouterr().out
    assert "数据不全" in out and "Ventas 源数据不完整" in out


@pytest.mark.parametrize("outcomes,unusable,per_store,expected", [
    (["ok", "ok"], [], True, 0),
    (["ok", "checks_failed"], [], True, 1),
    (["ok", "formula_errors"], [], True, 1),
    (["ok", "data_incomplete"], [], True, 1),
    (["checks_failed", "strict_aborted"], [], True, 2),
    (["ok", "timeout"], [], True, 2),
    (["ok", "stale"], [], True, 2),
    (["ok"], [("B", "why")], True, 2),
    (["ok"], [("B", "why")], False, 1),
])
def test_exit_code(outcomes, unusable, per_store, expected):
    results = [{"outcome": o} for o in outcomes]
    assert run_reports.exit_code(results, [("A", "x")], unusable, per_store) == expected


def test_exit_code_without_usable_store():
    assert run_reports.exit_code([], [], [("A", "why")]) == 2


# ------------------------------------------- financial_report.build（真子进程）

STUB = '''# -*- coding: utf-8 -*-
import io, json, sys
args = sys.argv[1:]
out = args[args.index("--out") + 1]
mode = %(mode)r
checks = [{"key": "bridge_b", "store": "A", "name": "b", "ok": mode == "clean",
           "severity": "error", "skipped": False, "detail": "d"}]
print("[REPORT_CHECKS_JSON] " + json.dumps(checks))
if mode == "crash":
    raise_it = {}["boom"]
if mode == "strict":
    sys.exit(10)
if mode == "crash70":
    print("Traceback (most recent call last):")
    sys.exit(70)
with io.open(out, "w") as f:
    f.write(u"x")
sys.exit(0 if mode == "clean" else 1)
'''


@pytest.fixture
def stub_build(monkeypatch, tmp_path):
    diag = tmp_path / "logs" / "report_diag"
    monkeypatch.setattr(financial_report, "DIAG_DIR", str(diag))

    def run(mode):
        script = tmp_path / ("stub_%s.py" % mode)
        script.write_text(STUB % {"mode": mode}, encoding="utf-8")
        monkeypatch.setattr(financial_report, "SCRIPT", str(script))
        out = tmp_path / "reports" / ("%s.xlsx" % mode)
        r = financial_report.build([("A", str(tmp_path))], "2026-08", str(out),
                                   no_recalc=True, echo=False)
        return r, out, diag

    return run


@pytest.mark.parametrize("mode,outcome", [("clean", "ok"), ("failed", "checks_failed")])
def test_written_report_leaves_no_diagnostics_file(stub_build, mode, outcome):
    r, out, diag = stub_build(mode)
    assert r["outcome"] == outcome and r["wrote"] and out.is_file()
    assert r["diagnostic_log"] is None
    assert not diag.exists()
    # 报表输出目录里只有报表：没有 diagnostics/ 子目录。
    assert os.listdir(str(out.parent)) == [out.name]


@pytest.mark.parametrize("mode,outcome", [
    ("strict", "strict_aborted"),
    ("crash70", "crashed"),
    ("crash", "not_written"),       # 解释器默认的 1：没有文件就不是"校验未过"
])
def test_unwritten_report_keeps_its_output_under_logs(stub_build, mode, outcome):
    r, out, diag = stub_build(mode)
    assert r["outcome"] == outcome and not r["wrote"] and not r["ok"]
    assert not financial_report.written(r)
    log = r["diagnostic_log"]
    assert log and os.path.dirname(log) == str(diag)
    assert not os.path.isdir(str(out.parent / "diagnostics"))
    with open(log, encoding="utf-8") as fh:
        text = fh.read()
    assert "2026-08" in text and "REPORT_CHECKS_JSON" in text
    if mode == "crash":
        assert "KeyError" in text       # traceback 留在现场记录里
