# -*- coding: utf-8 -*-
"""download_diag：归因规则、应有报表、取证、以及店内诊断这一步。全程不碰浏览器。"""
import argparse
import datetime
import json
import os

import pytest

import download_diag as dd
import meli_forms
import mercadopago
import run_downloads
import run_report

T0 = 1_800_000_000.0        # 任意一个固定时刻，规则不读时钟


def code(ev, **ctx):
    return dd.classify(ev, ctx)["cause"]


# ------------------------------------------------------------------ 归因规则

RULES = [
    # (说明, 证据, 上下文, 期望的原因码)
    ("过期", {"signal": "expired"}, {}, "EXPIRED"),
    ("店铺没打开", {"signal": "store-open", "exception_class": "ZiniaoError",
                "exception": "startBrowser failed: {...}", "step": "startBrowser"},
     {}, "STORE_OPEN_FAILED"),
    ("浏览器中途断开", {"signal": "exception", "exception_class": "WebDriverException",
                 "exception": "invalid session id"}, {}, "STORE_OPEN_FAILED"),
    ("登录页（URL）", {"signal": "ui-missing",
                  "url": "https://www.mercadolibre.com.mx/jms/mlm/lgz/login?go=x"},
     {}, "PAGE_BLOCKED"),
    ("验证页（标题）", {"signal": "ui-missing", "url": "https://x/y",
                 "title": "Verifica tu identidad"}, {}, "PAGE_BLOCKED"),
    ("浏览器错误页", {"signal": "exception", "exception_class": "RuntimeError",
                "exception": "Reportes tab not found",
                "url": "chrome-error://chromewebdata/"}, {}, "PAGE_BLOCKED"),
    ("权限墙（页面标记）", {"signal": "ui-missing", "permission_wall": True},
     {}, "NO_PERMISSION"),
    ("权限墙（正文那句话）", {"signal": "request-none",
                    "body": "Hola. No tienes permisos para ver los reportes."},
     {}, "NO_PERMISSION"),
    ("权限墙（异常）", {"signal": "no-permission",
                 "exception_class": "NoReportAccess"}, {}, "NO_PERMISSION"),
    ("无流水", {"row_state": "empty"}, {}, "NO_DATA"),
    ("无流水，哪怕语言不对", {"row_state": "empty", "lang": "en-US"}, {}, "NO_DATA"),
    ("数据处理中", {"row_state": "delayed"}, {}, "PLATFORM_DELAYED"),
    ("空文件", {"signal": "file-rejected", "files": ["a.csv（0 字节）"]},
     {}, "FILE_REJECTED"),
    ("账期没开始", {"signal": "period-not-started"}, {}, "NOT_AVAILABLE"),
    ("平台只给更少期", {"signal": "not-offered", "step": "仓储费只有 1 个账期可选"},
     {}, "NOT_AVAILABLE"),
    ("店铺没有 Full", {"signal": "ui-missing", "step": "库存报表菜单始终为空"},
     {"form": "stock/general", "has_full": False}, "NOT_AVAILABLE"),
    ("没有 Full 只管库存路线", {"signal": "ui-missing"},
     {"form": "billing/1", "has_full": False}, "UI_NOT_FOUND"),
    ("后续账期被重定向", {"signal": "wrong-period-page", "shown": "Septiembre",
                  "want": "Octubre"}, {"required": False}, "NOT_AVAILABLE"),
    ("目标账期被重定向", {"signal": "wrong-period-page", "shown": "Septiembre",
                  "want": "Julio"}, {"required": True}, "PAGE_BLOCKED"),
    ("就绪但文件没落盘", {"signal": "download-not-landed", "row_state": "ready"},
     {}, "DOWNLOAD_NOT_LANDED"),
    ("点了没文件，提示条说没有数据",
     {"signal": "download-not-landed",
      "alerts": ["No hay devoluciones en este período"]}, {}, "NO_DATA"),
    ("正文里的 no hay 不算数",
     {"signal": "download-not-landed", "body": "No hay datos", "alerts": []},
     {}, "DOWNLOAD_NOT_LANDED"),
    ("下载写了一半", {"signal": "download-not-landed",
                "crdownload": ["x.csv.crdownload"]}, {}, "DOWNLOAD_NOT_LANDED"),
    ("勾了 4 张到了 3 个", {"signal": "short-download", "checked": [1, 2, 3, 4],
                    "missing_reports": ["FULL"]}, {}, "DOWNLOAD_NOT_LANDED"),
    ("语言不对", {"signal": "ui-missing", "lang": "en-US"}, {}, "LANGUAGE_WRONG"),
    ("语言切不回去", {"signal": "exception", "exception_class": "RuntimeError",
                "exception": "界面语言切换后仍是 en_US（期望 es_MX）"},
     {}, "LANGUAGE_WRONG"),
    ("西语页面上控件缺失", {"signal": "ui-missing", "lang": "es-MX",
                   "step": "未找到 'Crear reporte'"}, {}, "UI_NOT_FOUND"),
    ("控件被禁用", {"signal": "ui-disabled", "step": "Generar 不可点击"},
     {}, "UI_NOT_FOUND"),
    ("selenium 找不到元素", {"signal": "exception",
                       "exception_class": "NoSuchElementException",
                       "exception": "no such element"}, {}, "UI_NOT_FOUND"),
    ("面板没渲染", {"signal": "exception", "exception_class": "RuntimeError",
               "exception": "Reportes panel never rendered"}, {}, "UI_NOT_FOUND"),
    ("脚本自己的异常", {"signal": "exception", "exception_class": "KeyError",
                 "exception": "'range'"}, {}, "INTERNAL_ERROR"),
    ("明细页没有月份", {"signal": "no-month-on-page"}, {}, "PAGE_BLOCKED"),
    ("就绪但没有下载入口", {"row_state": "ready", "fetch": "no-click-target"},
     {}, "UI_NOT_FOUND"),
    ("折叠面板：generar", {"row_state": "accordion", "action": "generar"},
     {}, "REQUEST_NOT_REGISTERED"),
    ("折叠面板：no-format-row", {"row_state": "accordion", "action": "no-format-row"},
     {}, "REQUEST_NOT_REGISTERED"),
    ("折叠面板：no-panel", {"row_state": "accordion", "action": "no-panel"},
     {}, "REQUEST_NOT_REGISTERED"),
    ("折叠面板：行在但没有链接", {"row_state": "accordion", "action": "no-link"},
     {}, "STILL_GENERATING"),
    ("折叠面板：没见过的状态", {"row_state": "accordion", "action": "reintentar"},
     {}, "UNKNOWN"),
    ("行在，还没好", {"row_state": "generating", "status_text": "En preparación"},
     {}, "STILL_GENERATING"),
    ("行不在：两次、超过 15 分钟", {"row_state": "absent"},
     {"requested_at": T0, "last_seen": T0 + 16 * 60, "absent_probes": 2},
     "REQUEST_NOT_REGISTERED"),
    ("行不在：只探测过一次", {"row_state": "absent"},
     {"requested_at": T0, "last_seen": T0 + 3600, "absent_probes": 1}, "UNKNOWN"),
    ("行不在：两次但才 5 分钟", {"row_state": "absent"},
     {"requested_at": T0, "last_seen": T0 + 300, "absent_probes": 2}, "UNKNOWN"),
    ("折叠面板没出现：两次、超过 15 分钟", {"row_state": "accordion", "action": None},
     {"requested_at": T0, "last_seen": T0 + 1000, "absent_probes": 3},
     "REQUEST_NOT_REGISTERED"),
    ("Ventas：最上面一条在生成", {"row_state": "ventas-widget", "widget_rows": 3,
                         "rows_before": 2, "top_sub": "Generando el archivo",
                         "top_has_dl": False}, {}, "STILL_GENERATING"),
    ("Ventas：成功记录但行数没涨", {"row_state": "ventas-widget", "widget_rows": 5,
                          "rows_before": 5, "top_sub": "Se generó con éxito",
                          "top_has_dl": True}, {}, "UNKNOWN"),
    ("Ventas：小组件是空的，够久", {"row_state": "ventas-widget", "widget_rows": 0},
     {"requested_at": T0, "last_seen": T0 + 2000, "absent_probes": 2},
     "REQUEST_NOT_REGISTERED"),
    ("平台 503 错误页（2026-10-13 实测）",
     {"url": "https://www.mercadopago.com.mx/balance/reports/account_statement",
      "title": "503 Service Temporarily Unavailable"}, {}, "PLATFORM_ERROR"),
    ("账单：少的是脚本不认识的报表类型（2026-10-13 EWTTO_SM 实测）",
     {"signal": "short-download", "checked": ["BILL_ML", "BILL_MP_V3", "FULL"],
      "missing_reports": [], "unrecognised_ids": ["BILL_MP_V3"]}, {},
     "UI_NOT_FOUND"),
    ("申请函数空手而回", {"signal": "request-none"}, {}, "UNKNOWN"),
    ("什么证据都没有", {}, {}, "UNKNOWN"),
]


@pytest.mark.parametrize("why,ev,ctx,want", RULES, ids=[r[0] for r in RULES])
def test_classify_rule(why, ev, ctx, want):
    rec = dd.classify(ev, ctx)
    assert rec["cause"] == want
    explain, owner, retry, action = dd.CAUSES[want]
    assert rec["explain"].startswith(explain)
    assert (rec["owner"], rec["retry"], rec["action"]) == (owner, retry, action)


def test_every_cause_code_is_reachable_and_documented():
    produced = {r[3] for r in RULES}
    assert produced == set(dd.CAUSES)
    for explain, owner, retry, action in dd.CAUSES.values():
        assert explain and action and isinstance(retry, bool)
        assert owner in ("平台", "账号", "脚本", "紫鸟")


def test_unknown_keeps_its_evidence_and_says_it_cannot_tell():
    ev = {"row_state": "absent", "url": "https://mp/reports"}
    rec = dd.classify(ev, {"form": "mercadopago/release", "requested_at": T0,
                           "last_seen": T0 + 60, "absent_probes": 1})
    assert rec["cause"] == "UNKNOWN"
    assert rec["evidence"]["url"] == "https://mp/reports"
    assert "分不清" in rec["explain"]


def test_record_shape():
    rec = dd.classify({"row_state": "generating"},
                      {"form": "mercadopago/collection", "period": "1/sep a 1/oct",
                       "first_seen": T0, "last_seen": T0 + 40, "probes": 3})
    assert set(rec) == {"form", "period", "cause", "explain", "owner", "retry",
                        "action", "evidence", "first_seen", "last_seen", "probes"}
    assert rec["form"] == "mercadopago/collection"
    assert rec["probes"] == 3
    assert rec["first_seen"] < rec["last_seen"]
    json.dumps(rec)                       # 要能原样进 JSON


def test_classify_is_pure_and_survives_garbage():
    ev = {"row_state": "accordion", "action": 7}     # 类型不对
    before = dict(ev)
    assert dd.classify(ev, None)["cause"] == "UNKNOWN"
    assert ev == before


# ------------------------------------------------------------------ 应有的报表

def args_for(*argv):
    return run_downloads.add_route_args(argparse.ArgumentParser()).parse_args(
        list(argv))


def test_expected_forms_default_route():
    got = dd.expected_forms(args_for(), today=datetime.date(2026, 10, 13))
    assert got == (["ventas"]
                   + ["mercadopago/%s" % k for k in mercadopago.REPORTS]
                   + ["billing/1", "billing/2"]
                   + ["stock/costos/1", "stock/costos/2", "stock/general",
                      "stock/consolidado", "stock/devoluciones"])


def test_expected_forms_month_route_stops_at_periods_not_started():
    a = args_for("--month", "2026-08")
    # 9 月中旬：8 月 + 9 月，10 月还没开始
    assert [f for f in dd.expected_forms(a, today=datetime.date(2026, 9, 15))
            if f.startswith("billing/")] == ["billing/2026-08", "billing/2026-09"]
    # 年底跨年
    a = args_for("--month", "2026-11", "--billing-after", "3")
    assert [f for f in dd.expected_forms(a, today=datetime.date(2027, 3, 1))
            if f.startswith("billing/")] == [
        "billing/2026-11", "billing/2026-12", "billing/2027-01", "billing/2027-02"]


def test_expected_forms_honours_only_and_months():
    assert dd.expected_forms(args_for("--only", "ventas")) == ["ventas"]
    assert dd.expected_forms(args_for("--only", "billing", "--months", "3")) == [
        "billing/1", "billing/2", "billing/3"]
    assert dd.expected_forms(args_for("--only", "stock")) == list(dd.STOCK_FORMS)


def test_route_lists_match_the_page_driving_modules():
    # download_diag 抄了一份清单以免 import 页面驱动层；这里盯着两边一致
    assert list(dd.MP_REPORTS) == list(mercadopago.REPORTS)
    keys = [k for k, _ in meli_forms.STOCK_REPORTS]
    assert {f.split("/")[1] for f in dd.STOCK_FORMS} == set(keys)
    assert dd.PERMISSION_WALL == mercadopago.PERMISSION_WALL


# ------------------------------------------------------------------ 取证

class FakeDriver(object):
    current_url = "https://www.mercadopago.com.mx/balance/reports/release"
    title = "Reportes"

    def __init__(self, facts=None):
        self.facts = facts or {"lang": "es-MX", "body": "Hola " * 100,
                               "alerts": ["Algo salió mal"],
                               "permission_wall": False}
        self.shots = []

    def execute_script(self, js, *a):
        assert js == dd.JS_PAGE_FACTS      # 只读那一段，不点任何东西
        return self.facts

    def save_screenshot(self, path):
        self.shots.append(path)
        with open(path, "wb") as f:
            f.write(b"png")
        return True


class BrokenDriver(object):
    @property
    def current_url(self):
        raise RuntimeError("invalid session id")

    @property
    def title(self):
        raise RuntimeError("invalid session id")

    def execute_script(self, *a):
        raise RuntimeError("invalid session id")

    def save_screenshot(self, path):
        raise RuntimeError("invalid session id")


def test_capture_evidence_saves_page_facts_and_screenshot(tmp_path):
    d = FakeDriver()
    ev = dd.capture_evidence(d, "mercadopago/release", str(tmp_path),
                             extra={"signal": "ui-missing", "step": "x",
                                    "menu_items": ["a", "b"], "obj": object()})
    assert ev["url"] == d.current_url and ev["title"] == "Reportes"
    assert ev["lang"] == "es-MX" and ev["alerts"] == ["Algo salió mal"]
    assert ev["signal"] == "ui-missing" and ev["menu_items"] == ["a", "b"]
    assert ev["screenshot"].startswith("_diag/") and ev["screenshot"].endswith(".png")
    assert os.path.isfile(os.path.join(str(tmp_path), ev["screenshot"]))
    saved = json.load(open(os.path.join(str(tmp_path), ev["file"]), encoding="utf-8"))
    assert saved["step"] == "x" and saved["url"] == d.current_url
    json.dumps(ev)


def test_capture_evidence_never_raises(tmp_path):
    ev = dd.capture_evidence(BrokenDriver(), "ventas", str(tmp_path),
                             extra={"signal": "exception"})
    assert ev["url"] is None and ev["screenshot"] is None
    assert len(ev["capture_errors"]) >= 3
    assert ev["signal"] == "exception"
    # 连目录都写不了也不抛
    blocker = tmp_path / "file"
    blocker.write_text("x")
    ev = dd.capture_evidence(FakeDriver(), "ventas", str(blocker))
    assert ev["file"] is None and ev["url"]
    # 一个根本不是 driver 的东西
    assert dd.capture_evidence({"kind": "ventas"}, "x", str(tmp_path))["url"] is None


def test_capture_evidence_without_driver_or_screenshot(tmp_path):
    ev = dd.capture_evidence(None, "store", str(tmp_path),
                             extra={"signal": "store-open"})
    assert ev["no_driver"] is True and ev["screenshot"] is None and ev["file"]
    d = FakeDriver()
    ev = dd.capture_evidence(d, "poll", str(tmp_path / "other"), shot=False)
    assert ev["file"] is None and d.shots == []
    assert not (tmp_path / "other").exists()         # 轮询不落盘


def test_page_reading_permission_wall_feeds_the_rule(tmp_path):
    d = FakeDriver({"lang": "es-MX", "body": "nav", "alerts": [],
                    "permission_wall": True})
    ev = dd.capture_evidence(d, "mercadopago", str(tmp_path), shot=False)
    assert dd.classify(ev)["cause"] == "NO_PERMISSION"


# ------------------------------------------------------------------ 小工具

def test_checked_report_ids_are_deduplicated_in_order():
    # 同一个复选框被每个 root 各收一遍：4 张报表看起来像勾了 8 张
    once = [{"id": "«r2»", "checked": True}, {"id": "BILL_ML", "checked": True},
            {"id": "FULL", "checked": True}, {"id": "PAYMENT", "checked": False},
            {"id": "NC_ML", "checked": True}, {"id": "", "checked": True}]
    assert meli_forms.checked_report_ids(once + once, "«r2»") == [
        "BILL_ML", "FULL", "NC_ML"]
    assert meli_forms.checked_report_ids(None) == []


def test_match_billing_files_names_the_missing_report():
    files = ["Reporte_Facturacion_MercadoLibre_Ago2026.xlsx",
             "Reporte_Pagos_Facturas_Ago2026.xlsx",
             "Reporte_Notas_Credito_MercadoLibre_Ago2026.xlsx"]
    got, missing, unknown = dd.match_billing_files(
        ["BILL_ML", "FULL", "PAYMENT", "NC_ML", "NEW_ONE"], files)
    assert missing == ["FULL"] and unknown == ["NEW_ONE"]
    assert got["PAYMENT"] == files[1]


def test_folder_delta_separates_half_written_files():
    d = dd.folder_delta({"old.csv": 1},
                        {"old.csv": 1, "new.csv": 5, "x.csv.crdownload": 0,
                         "_diag": 0})
    assert d == {"folder_delta": ["new.csv"], "crdownload": ["x.csv.crdownload"]}


def test_snapshot_dir_ignores_the_evidence_folder(tmp_path):
    (tmp_path / "a.csv").write_text("x")
    os.makedirs(str(tmp_path / "_diag"))
    assert list(meli_forms.snapshot_dir(str(tmp_path))) == ["a.csv"]


# ------------------------------------------------------- 账单：勾选数不再翻倍

class Box(object):
    def get_attribute(self, name):
        return {"id": "«r2»", "class": "andes-tab--selected"}.get(name)

    def is_selected(self):
        return True

    def is_enabled(self):
        return True

    def click(self):
        pass


class BillingDriver(FakeDriver):
    def execute_script(self, js, *a):
        if js == meli_forms.JS_DEEP_CHECKBOX_STATE:
            once = [{"id": "«r2»", "checked": True}] + [
                {"id": i, "checked": True}
                for i in ("BILL_ML", "FULL", "PAYMENT", "NC_ML")]
            return once + once
        if js in (meli_forms.JS_DEEP_QUERY, meli_forms.JS_DEEP_CHECKBOX_BY_LABEL):
            return Box()
        if js == dd.JS_PAGE_FACTS:
            return self.facts
        return None


def run_detail(monkeypatch, tmp_path, files):
    seen = {}

    def fake_wait(path, before, min_files=1, **kw):
        seen["min_files"] = min_files
        return files

    monkeypatch.setattr(meli_forms.time, "sleep", lambda s: None)
    monkeypatch.setattr(meli_forms, "wait_for_downloads", fake_wait)
    diag = dd.Collector(str(tmp_path), "S")
    got = meli_forms._download_detail_reports(
        BillingDriver(), str(tmp_path), "Agosto", diag=diag, form="billing/1",
        period="Agosto (20260831)")
    return got, seen, diag


def test_billing_expects_four_files_not_eight(monkeypatch, tmp_path, capsys):
    files = ["Reporte_Facturacion_MercadoLibre_Ago2026.xlsx",
             "Reporte_Cargos_Full_Ago2026.xlsx",
             "Reporte_Pagos_Facturas_Ago2026.xlsx",
             "Reporte_Notas_Credito_MercadoLibre_Ago2026.xlsx"]
    got, seen, diag = run_detail(monkeypatch, tmp_path, files)
    assert got == files and seen["min_files"] == 4
    out = capsys.readouterr().out
    assert "已勾选报表（4）" in out and "[警告]" not in out
    assert "billing/1" in diag.delivered and not diag.items


def test_billing_shortfall_names_the_missing_report(monkeypatch, tmp_path, capsys):
    files = ["Reporte_Facturacion_MercadoLibre_Ago2026.xlsx",
             "Reporte_Pagos_Facturas_Ago2026.xlsx",
             "Reporte_Notas_Credito_MercadoLibre_Ago2026.xlsx"]
    got, _seen, diag = run_detail(monkeypatch, tmp_path, files)
    assert "缺的是：FULL" in capsys.readouterr().out
    assert list(diag.items) == ["billing/1/FULL"]
    recs = dd.diagnose_store(["billing/1"], diag)
    assert [(r["form"], r["cause"]) for r in recs] == [
        ("billing/1/FULL", "DOWNLOAD_NOT_LANDED")]
    assert "FULL" in recs[0]["explain"]


def test_billing_nothing_landed_is_recorded(monkeypatch, tmp_path):
    got, _seen, diag = run_detail(monkeypatch, tmp_path, [])
    assert got == []
    recs = dd.diagnose_store(["billing/1"], diag)
    assert [(r["form"], r["cause"], r["period"]) for r in recs] == [
        ("billing/1", "DOWNLOAD_NOT_LANDED", "Agosto (20260831)")]


# ------------------------------------------------------------------ 收集器

def test_collector_keeps_the_first_scene_and_merges_flags(tmp_path):
    c = dd.Collector(str(tmp_path), "S")
    c.give_up(None, "stock/consolidado", "ui-disabled", step="Descargar 不可点击",
              shot=False)
    c.give_up(None, "stock/consolidado", "download-not-landed", shot=False)
    ev = c.items["stock/consolidado"]["evidence"]
    assert ev["signal"] == "ui-disabled"
    assert ev["later"] == [{"signal": "download-not-landed", "step": None}]
    c.flag("range_clamped", {"ventas": {"requested_months": 6, "actual_months": 2}})
    c.flag("range_clamped", {"stock/consolidado": {"actual_start": "2026-08-01"}})
    assert set(c.flags["range_clamped"]) == {"ventas", "stock/consolidado"}


def test_module_level_helpers_are_noops_without_a_collector():
    assert dd.give_up(None, FakeDriver(), "ventas", "ui-missing") is None
    dd.delivered(None, "ventas", ["a"])
    dd.flag(None, "locale", "en_US")


def test_give_up_helper_never_raises(tmp_path, capsys):
    class Bad(object):
        def give_up(self, *a, **k):
            raise ValueError("boom")
    assert dd.give_up(Bad(), None, "ventas", "ui-missing") is None
    assert "诊断取证失败" in capsys.readouterr().out


# ------------------------------------------------------------------ 店内诊断

def test_diagnose_store_computes_what_is_missing(tmp_path):
    out = str(tmp_path)
    (tmp_path / "ok.xlsx").write_text("data")
    (tmp_path / "empty.csv").write_text("")
    c = dd.Collector(out, "S")
    c.mark_delivered("billing/1", ["ok.xlsx"], "Agosto (20260831)")
    c.mark_delivered("stock/general", ["empty.csv"])
    c.give_up(None, "stock/devoluciones", "ui-missing", shot=False,
              step="菜单里没有 'reporte de devoluciones a la bodega'")
    c.give_up(None, "mercadopago", "no-permission", shot=False)
    c.give_up(None, "billing", "exception", shot=False,
              exception_class="KeyError", exception="'x'")
    owed = {"kind": "mercadopago", "report": "release", "range": "1/ago a 1/oct"}
    dd.note_probe(owed, {"row_state": "generating", "status_text": "En preparación"},
                  now=T0)
    nodata = {"kind": "mercadopago", "report": "withdraw", "range": "r"}
    dd.note_probe(nodata, {"row_state": "empty"}, now=T0)
    expected = ["ventas", "mercadopago/release", "mercadopago/withdraw",
                "mercadopago/collection", "billing/1", "billing/2",
                "stock/general", "stock/devoluciones"]
    recs = dd.diagnose_store(expected, c, entries=[owed], no_data=[nodata],
                             collected={"ventas": ["ok.xlsx"]}, out_dir=out)
    got = {r["form"]: r for r in recs}
    assert "ventas" not in got and "billing/1" not in got
    assert got["mercadopago/release"]["cause"] == "STILL_GENERATING"
    assert got["mercadopago/release"]["owed"] is True
    assert got["mercadopago/release"]["period"] == "1/ago a 1/oct"
    assert got["mercadopago/withdraw"]["cause"] == "NO_DATA"
    assert "owed" not in got["mercadopago/withdraw"]
    # 没来得及申请的那张，跟着整条路线的原因走
    assert got["mercadopago/collection"]["cause"] == "NO_PERMISSION"
    assert got["mercadopago/collection"]["evidence"]["inherited_from"] == "mercadopago"
    assert got["billing/2"]["cause"] == "INTERNAL_ERROR"
    assert got["stock/general"]["cause"] == "FILE_REJECTED"
    assert got["stock/devoluciones"]["cause"] == "UI_NOT_FOUND"


def test_diagnose_store_without_any_evidence_is_unknown_not_silent():
    recs = dd.diagnose_store(["stock/general"], dd.Collector("/nonexistent"))
    assert recs[0]["cause"] == "UNKNOWN" and "没有采到任何证据" in recs[0]["explain"]


def test_periods_not_attempted_inherit_the_cause_that_stopped_the_loop(tmp_path):
    c = dd.Collector(str(tmp_path))
    c.give_up(None, "billing/2026-09", "wrong-period-page", shot=False,
              required=False, shown="Agosto", want="Septiembre")
    c.give_up(None, "billing/2026-10", "not-attempted", shot=False,
              after="billing/2026-09")
    recs = dd.diagnose_store(["billing/2026-08", "billing/2026-09",
                              "billing/2026-10"], c,
                             collected={"billing/2026-08": []})
    got = {r["form"]: r for r in recs}
    assert got["billing/2026-09"]["cause"] == "NOT_AVAILABLE"
    assert got["billing/2026-10"]["cause"] == "NOT_AVAILABLE"
    assert "billing/2026-09" in got["billing/2026-10"]["explain"]


def test_route_errors_only_for_missing_that_needs_attention():
    recs = [dd.classify({"signal": "ui-missing"}, {"form": "stock/general"}),
            dd.classify({"signal": "not-offered"}, {"form": "stock/costos/2"}),
            dd.classify({"signal": "wrong-period-page"},
                        {"form": "billing/2026-07", "required": True}),
            dd.classify({"row_state": "generating"}, {"form": "mercadopago/release"})]
    errs = run_downloads.route_errors(recs)
    assert len(errs) == 2
    assert errs[0].startswith("billing: 1 张未下载") and "PAGE_BLOCKED" in errs[0]
    assert "stock/general [UI_NOT_FOUND]" in errs[1] and "costos" not in errs[1]


# ------------------------------------------------------------------ 待取条目

def test_note_probe_keeps_only_the_latest_state_plus_counters():
    e = {"kind": "mercadopago", "report": "settlement", "range": "r",
         "requested_at": T0}
    dd.note_probe(e, {"row_state": "absent"}, now=T0 + 60)
    dd.note_probe(e, {"row_state": "absent"}, now=T0 + 1000)
    assert (e["probes"], e["absent_probes"]) == (2, 2)
    assert (e["first_probe_at"], e["last_probe_at"]) == (T0 + 60, T0 + 1000)
    assert e["last_state"] == "absent"
    assert dd.record_from_entry(e)["cause"] == "REQUEST_NOT_REGISTERED"
    # 行出现了：之前的"不在"作废，旧字段不留
    dd.note_probe(e, {"row_state": "generating", "status_text": "En preparación"},
                  now=T0 + 1100)
    assert e["absent_probes"] == 0 and e["last_status_text"] == "En preparación"
    rec = dd.record_from_entry(e)
    assert rec["cause"] == "STILL_GENERATING" and rec["probes"] == 3
    json.dumps(e)


def test_note_failure_marks_store_open_and_discards_stale_row_state():
    class ZiniaoError(RuntimeError):
        pass
    e = {"kind": "ventas", "requested_at": T0}
    dd.note_probe(e, {"row_state": "ventas-widget", "widget_rows": 2}, now=T0)
    err = ZiniaoError("startBrowser failed: {\"statusCode\": -10003}")
    err.ziniao_step, err.ziniao_response = "startBrowser", '{"statusCode": -10003}'
    dd.note_failure(e, err)
    rec = dd.record_from_entry(e)
    assert rec["cause"] == "STORE_OPEN_FAILED" and rec["form"] == "ventas"
    assert "startBrowser" in rec["explain"] and "-10003" in rec["explain"]
    assert e["probes"] == 1                 # 没看到页面，不算一次探测
    dd.note_failure(e, KeyError("range"))
    assert dd.record_from_entry(e)["cause"] == "INTERNAL_ERROR"


def test_store_failure_records_cover_every_expected_form():
    recs = dd.store_failure_records(
        type("ZiniaoError", (RuntimeError,), {})("store 'X' not found"),
        ["ventas", "billing/1"])
    assert [(r["form"], r["cause"]) for r in recs] == [
        ("ventas", "STORE_OPEN_FAILED"), ("billing/1", "STORE_OPEN_FAILED")]


def test_merge_final_replaces_owed_verdicts_and_keeps_the_rest():
    sync = dd.classify({"signal": "ui-missing"}, {"form": "stock/general"})
    failed_request = dd.classify({"signal": "ui-disabled"},
                                 {"form": "mercadopago/collection"})
    was_owed = dict(dd.classify({"row_state": "absent"},
                                {"form": "mercadopago/release", "period": "r"}),
                    owed=True)
    collected_late = dict(dd.classify({"row_state": "generating"}, {"form": "ventas"}),
                          owed=True)
    still = {"kind": "mercadopago", "report": "release", "range": "r",
             "requested_at": T0}
    dd.note_probe(still, {"row_state": "absent"}, now=T0 + 600)
    dd.note_probe(still, {"row_state": "absent"}, now=T0 + 4000)
    empty = {"kind": "mercadopago", "report": "withdraw", "range": "w"}
    dd.note_probe(empty, {"row_state": "empty"}, now=T0)
    final = dd.merge_final([sync, failed_request, was_owed, collected_late],
                           [still], [empty])
    got = {r["form"]: r["cause"] for r in final}
    assert got == {"stock/general": "UI_NOT_FOUND",
                   "mercadopago/collection": "UI_NOT_FOUND",
                   # 店内只能说分不清，补收之后分得清了
                   "mercadopago/release": "REQUEST_NOT_REGISTERED",
                   "mercadopago/withdraw": "NO_DATA"}
    assert dd.merge_final(None) == []


# ------------------------------------------------------------------ 输出

def sample_rows():
    a = dd.classify({"signal": "store-open", "exception": "boom"}, {"form": "ventas"})
    b = dd.classify({"signal": "store-open", "exception": "boom"},
                    {"form": "billing/1"})
    c = dd.classify({"row_state": "empty"},
                    {"form": "mercadopago/withdraw", "period": "r"})
    return [{"store": "A", "missing": [a, b]},
            {"store": "B", "missing": [c],
             "cautions": [{"form": "mercadopago/after_collection",
                           "requested": "2026-08-14 .. 2026-10-13",
                           "adopted": "4 julio 2026 - 2 septiembre 2026"}]},
            {"store": "C", "missing": []}]


def test_format_section_lists_store_form_cause_retry_action_and_counts():
    text = "\n".join(dd.format_section(sample_rows()))
    assert "未下载的报表与原因" in text
    assert "A | ventas、billing/1" in text           # 同店同因并成一行
    assert "[STORE_OPEN_FAILED]" in text and "重试有用" in text
    assert "下一步  " + dd.CAUSES["STORE_OPEN_FAILED"][3] in text
    assert "B | mercadopago/withdraw（r）" in text and "重试没用" in text
    assert "按原因统计" in text and "STORE_OPEN_FAILED" in text
    assert "NO_DATA" in text and "（无需处理）" in text
    assert "申请 2026-08-14 .. 2026-10-13 → 采用 4 julio 2026" in text


def test_format_section_says_so_when_nothing_is_missing():
    text = "\n".join(dd.format_section([{"store": "A", "missing": []}]))
    assert "全部到齐" in text and "按原因统计" not in text


def result_for(tmp_path, missing):
    return {"store": "S", "out_dir": str(tmp_path), "billing": {}, "sales": [],
            "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
            "seconds": 5, "missing": missing,
            "cautions": [{"form": "mercadopago/after_collection",
                          "requested": "a|b", "adopted": "c", "reason": "顶行没变"}],
            "flags": {"ip_check_failed": True, "locale": "en_US",
                      "range_clamped": {"ventas": {"requested_months": 6,
                                                   "actual_months": 2}}}}


def test_report_md_has_the_why_missing_table(tmp_path):
    rec = dd.classify({"signal": "ui-missing", "file": "_diag/1_ventas.json"},
                      {"form": "ventas", "period": "últimos 6 meses"})
    text = run_report.render(result_for(tmp_path, [rec]))
    assert "## 未下载 —— 原因" in text
    assert "| 报表 | 期间 | 原因 | 重试有用? | 下一步 | 证据 |" in text
    row = next(l for l in text.splitlines() if l.startswith("| `ventas`"))
    assert "**UI_NOT_FOUND**" in row and "| 否 |" in row
    assert "`_diag/1_ventas.json`" in row and "últimos 6 meses" in row
    assert "a/b" in text                          # 竖线不能撑破表格
    assert "ip_check_failed" in text and "range_clamped" in text and "en_US" in text
    assert text.index(run_report.MISSING_BEGIN) < text.index(run_report.MISSING_END)
    # 没做过诊断的结果（download_history）不出这一段
    assert "未下载" not in run_report.render(
        {k: v for k, v in result_for(tmp_path, []).items() if k != "missing"})


def test_update_missing_replaces_only_that_section(tmp_path):
    rec = dd.classify({"row_state": "generating"}, {"form": "mercadopago/release"})
    res = result_for(tmp_path, [rec])
    path = run_report.write(res)
    before = open(path, encoding="utf-8").read()
    assert "STILL_GENERATING" in before
    run_report.update_missing(str(tmp_path), [], late_files=["release-1.csv"])
    after = open(path, encoding="utf-8").read()
    assert "STILL_GENERATING" not in after and "全部到齐" in after
    assert "`release-1.csv`" in after
    assert after.count(run_report.MISSING_BEGIN) == 1
    assert after.startswith(before[:before.index(run_report.MISSING_BEGIN)])
    assert after.rstrip().endswith("`FORMS.md`.")
    # 老报告没有记号：插在页脚之前
    with open(path, "w", encoding="utf-8") as f:
        f.write("# old\n\nbody\n\n---\n\nfooter\n")
    run_report.update_missing(str(tmp_path), [rec])
    old = open(path, encoding="utf-8").read()
    assert old.index("STILL_GENERATING") < old.index("footer")
    assert run_report.update_missing(str(tmp_path / "nope"), []) is None


def test_select_all_that_did_not_take_is_retried(monkeypatch, tmp_path, capsys):
    """第一次点全选时表格还没就绪：一个报表都没勾上（2026-10-14 TRQWH_TA01 9 月）。"""
    monkeypatch.setattr(meli_forms.time, "sleep", lambda s: None)
    monkeypatch.setattr(meli_forms, "wait_for_downloads", lambda *a, **k: ["a", "b", "c", "d"])
    real, calls = meli_forms.checked_report_ids, []

    def first_empty(state, select_all_id=None):
        calls.append(1)
        return [] if len(calls) == 1 else real(state, select_all_id)

    monkeypatch.setattr(meli_forms, "checked_report_ids", first_empty)
    got = meli_forms._download_detail_reports(BillingDriver(), str(tmp_path), "Septiembre",
                                              diag=dd.Collector(str(tmp_path), "S"))
    out = capsys.readouterr().out
    assert len(calls) == 2, "第一次没勾上，要再看一次"
    assert "全选没有生效（第 1 次）" in out and "已勾选报表（4）" in out
    assert got == ["a", "b", "c", "d"]
