# -*- coding: utf-8 -*-
"""The window's 期间报表 tab and the several-day schedule, built with the fake
tkinter (see test_app_smoke.py): no real process, no real task."""
import io
import os

import fake_tk
import period_data as P
from core import region_choice
from test_app_smoke import _sync_threads, app3  # noqa: F401  (fixture)


def test_schedule_with_several_days(app3, monkeypatch):  # noqa: F811
    from core import app as app_mod
    from core import entrance
    _sync_threads(monkeypatch)
    got = []
    monkeypatch.setattr(app_mod.W, "install", lambda name, xml: got.append((name, xml)) or (0, ""))
    app3._show_schedule({region_choice.ALL_TASK: None})
    app3.days.set("26, 3，12")
    fake_tk.ANSWERS[:] = [True]
    app3.do_install()
    app3.fake_root.pump()
    name, xml = got[-1]
    assert name == "月度报表" and "--region ALL" in xml           # same task, same arguments
    assert xml.count("<CalendarTrigger>") == 1 and "<Day>3</Day><Day>12</Day><Day>26</Day>" in xml
    assert "每月 3、12、26 号" in fake_tk.SHOWN[-1][2]
    s = entrance.load_settings()
    assert s["sched_days_ALL"] == [3, 12, 26] and s["sched_day_ALL"] == 3
    # a registered several-day task shows its days
    reg = {"DAY": "3,12,26", "DAYS": [3, 12, 26], "TIME": "02:00", "RESULT": "0", "XML": "",
           "ARGS": "--entrance scheduled --region ALL"}
    app3.days.set("15")
    app3._show_schedule({region_choice.ALL_TASK: reg})
    assert app3.days.get() == "3, 12, 26"
    assert app3.sched_head.get().startswith("● 已设置：每月 3、12、26 号 02:00")
    # a bad day is refused before anything is registered
    n = len(got)
    app3.days.set("3, 31")
    app3.do_install()
    assert len(got) == n and fake_tk.SHOWN[-1][0] == "showwarning" and "31" in fake_tk.SHOWN[-1][2]


def test_old_single_day_setting_is_shown(tmp_path, monkeypatch):
    import json
    from test_app_smoke import MULTI, _app
    (tmp_path / "gui_settings.json").write_text(json.dumps({"sched_day_ALL": 9}), encoding="utf-8")
    gen = _app(tmp_path, monkeypatch, MULTI)
    a = next(gen)
    try:
        assert a.days.get() == "9"
    finally:
        gen.close()


def test_period_tab_previews_and_generates(app3, tmp_path, monkeypatch):  # noqa: F811
    from core import app as app_mod
    out = P.configure(tmp_path, monkeypatch, v2=True)
    app3.outdir.set(str(tmp_path / "share"))                   # never the real default folder
    _sync_threads(monkeypatch)                                 # the preview reads in a thread
    for m in ("2026-07", "2026-08"):
        P.put("BOCINA_SM", m, "20261001_000000")
    P.put("BOCINA_SM", "2026-09", "20261005_000000", level="gaps", reasons=["Ventas 没下载到"])
    assert app3.tabs.tabs_[3][1]["text"].strip() == "期间报表"
    app3.show_tab("period")
    app3.period_kind.set("quarterly")
    app3.period_quarter.set("2026-Q3")
    app3._period_changed()
    tv = app3.period_grid
    assert tv.get_children("") == ("m:ML-MX",)
    assert tv.text("m:ML-MX") == "美客多 墨西哥（账单基准）"
    assert tv.item("s:ML-MX:BOCINA_SM")["values"] == ["✓ 完整", "✓ 完整", "✗ 无完整版本"]
    assert "1 个店铺月份没有完整版本" in app3.period_grid_var.get()
    tv.focus("s:ML-MX:BOCINA_SM")
    app3._period_reason()
    assert "BOCINA_SM 2026-09" in app3.period_grid_var.get() and "Ventas 没下载到" in app3.period_grid_var.get()
    assert app3.btn_period.cget("state") == "normal"
    assert not os.path.exists(os.path.join(out, "quarterly"))          # the preview writes nothing

    ran = []

    class Proc(object):
        def __init__(self, cmd, **kw):
            ran.append(cmd)
            self.stdout = io.BytesIO("期间报表：2026 年第 3 季度\n结果：已生成\n".encode("utf-8"))

        def wait(self):
            return 1
    monkeypatch.setattr(app_mod.subprocess, "Popen", Proc)
    _sync_threads(monkeypatch)
    app3.btn_period.invoke()
    app3.fake_root.pump()
    assert ran[-1][2].endswith("run_period.py")
    assert ran[-1][3:] == ["--quarter", "2026-Q3", "--region", "ALL", "--out-dir", app3._out_dir()]
    assert "结果：已生成" in app3.period_out.text()
    assert app3.period_status.get().startswith("⚠ 期间报表生成了")
    assert app3.period_proc is None and app3.btn_period.cget("state") == "normal"

    # one country, order basis: no order results -> nothing to generate
    app3.period_picker.set_all(False)
    app3.period_picker.toggle("u:MX")
    app3.period_basis.set("order")
    app3._period_changed()
    assert app3.period_grid.get_children("") == ()
    assert "还没有任何月度结果" in app3.period_grid_var.get()
    assert app3.btn_period.cget("state") == "disabled"

    # a custom period the wrong way round
    app3.period_basis.set("bill")
    app3.period_kind.set("custom")
    app3.period_from.set("2026-09")
    app3.period_to.set("2026-08")
    app3._period_changed()
    assert app3.period_text.get().startswith("✗") and app3.btn_period.cget("state") == "disabled"

    # open the folder: none yet -> a message; after a generation -> the market's folder
    opened = []
    monkeypatch.setattr(app_mod, "open_path", opened.append)
    app3.period_kind.set("quarterly")
    app3._period_changed()
    app3.do_open_period_dir()
    assert opened == [] and fake_tk.SHOWN[-1][1] == "还没有期间报表"
    from core import period_report as PR
    PR.generate(PR.quarter("2026-Q3"), "MX", out_dir=app3._out_dir())
    app3.do_open_period_dir()
    assert opened[-1].endswith(os.path.join("quarterly", "2026-Q3", "ML-MX"))


def test_period_preview_reads_off_the_tk_thread(app3, tmp_path, monkeypatch):  # noqa: F811
    import time
    P.configure(tmp_path, monkeypatch, v2=True)
    P.put("BOCINA_SM", "2026-07", "20261001_000000")
    app3.period_kind.set("quarterly")
    app3.period_quarter.set("2026-Q3")
    app3._period_changed()                         # a real thread; the answer comes through after()
    for _ in range(200):
        app3.fake_root.pump(1)
        if app3.period_grid.get_children(""):
            break
        time.sleep(0.01)
    assert app3.period_grid.get_children("") == ("m:ML-MX",)
