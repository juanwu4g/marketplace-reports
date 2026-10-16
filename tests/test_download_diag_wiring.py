# -*- coding: utf-8 -*-
"""下载层在"放弃"的每个出口都把现场交给了收集器，而且原来的返回值没变。

页面是假的：一个按 JS 文本查表回答的 driver。测的是接线，不是选择器 ——
选择器只有真店铺能验。
"""
import datetime

import pytest

from core import download_diag as dd
from platforms.mercadolibre import meli_forms as mf
from platforms.mercadolibre import mercadopago as mp


class Clock(object):
    def __init__(self):
        self.now = 1_800_000_000.0

    def time(self):
        return self.now

    def sleep(self, s):
        self.now += s


class El(object):
    text = "x"

    def click(self):
        pass


class Page(object):
    """execute_script 按脚本文本查表；值是可调用的就用参数调它。没列的返回 None。"""
    title = "Mercado Libre"

    def __init__(self, answers=None, url="https://vendedores.mercadolibre.com.mx/x",
                 buttons=()):
        self.answers = dict(answers or {})
        self.current_url = url
        self.buttons = list(buttons)
        self.calls = []
        self.shots = 0

    def get(self, url):
        self.current_url = url

    def refresh(self):
        pass

    def find_elements(self, by, sel):
        return self.buttons

    def execute_script(self, js, *a):
        self.calls.append(js)
        if js == dd.JS_PAGE_FACTS:
            return {"lang": "es-MX", "body": "", "alerts": []}
        v = self.answers.get(js)
        return v(*a) if callable(v) else v

    def save_screenshot(self, path):
        self.shots += 1
        return True


@pytest.fixture
def world(monkeypatch, tmp_path):
    clock = Clock()
    monkeypatch.setattr(mf, "time", clock)
    monkeypatch.setattr(mp, "time", clock)
    return dd.Collector(str(tmp_path), "S"), str(tmp_path)


def causes(diag, expected):
    return {r["form"]: r["cause"] for r in dd.diagnose_store(expected, diag)}


# ------------------------------------------------------------------ 库存

def test_stock_menu_never_opens(world):
    diag, out = world
    got = mf.download_stock_reports(Page({mf.JS_OPEN_STOCK_MENU: False}), out,
                                    diag=diag)
    assert got == {"costos": [], "general": [], "consolidado": [],
                   "devoluciones": []}                    # 返回值的形状没变
    assert causes(diag, dd.STOCK_FORMS) == {f: "UI_NOT_FOUND" for f in dd.STOCK_FORMS}
    assert diag.items["stock/general"]["evidence"]["menu_items"] == []


def test_stock_each_report_gets_its_own_reason(world, monkeypatch):
    diag, out = world
    # 第 4 个 [] 是 devoluciones 的重试：点了菜单没有下载开始，会重开页面再点一次
    arrivals = iter([["Costos_1.xlsx"], ["stock_general_full_1.xlsx"], [], [], []])
    monkeypatch.setattr(mf, "wait_for_downloads",
                        lambda *a, **k: next(arrivals))
    page = Page({mf.JS_OPEN_STOCK_MENU: True, mf.JS_STOCK_ITEMS: ["Reporte x"],
                 mf.JS_CLICK_STOCK_ITEM: True,
                 mf.JS_LIST_PERIODS: ["Del 1 de septiembre al 30 de septiembre del 2026"],
                 mf.JS_PICK_PERIOD: True,
                 mf.JS_MODAL_BUTTON: lambda key: (
                     {"found": True, "disabled": False} if key == "almacenamiento"
                     else {"found": True, "disabled": True}),
                 mf.JS_OPEN_DATEPICKER: True, mf.JS_CAL_CAPTION: "nunca 1999",
                 mf.JS_CAL_FIRST_ENABLED: None})
    got = mf.download_stock_reports(page, out, diag=diag)
    assert got == {"costos": ["Costos_1.xlsx"],
                   "general": ["stock_general_full_1.xlsx"],
                   "consolidado": [], "devoluciones": []}
    assert set(diag.delivered) == {"stock/costos/1", "stock/general"}
    assert diag.delivered["stock/costos/1"]["period"].startswith("Del 1 de septiembre")
    recs = {r["form"]: r for r in dd.diagnose_store(dd.STOCK_FORMS, diag)}
    assert {f: r["cause"] for f, r in recs.items()} == {
        "stock/costos/2": "NOT_AVAILABLE",          # 平台只给了一个账期
        "stock/consolidado": "UI_NOT_FOUND",        # 日期没选上，不是"没下到"
        "stock/devoluciones": "DOWNLOAD_NOT_LANDED"}
    assert "日期区间没有选上" in recs["stock/consolidado"]["explain"]
    assert mf.JS_CLICK_MODAL_BUTTON in page.calls        # 仓储费那一期照常点了
    assert next(arrivals, "done") == "done", "devoluciones 重试过一次"


def test_stock_without_collector_behaves_as_before(world, monkeypatch):
    _diag, out = world
    monkeypatch.setattr(mf, "wait_for_downloads", lambda *a, **k: [])
    got = mf.download_stock_reports(Page({mf.JS_OPEN_STOCK_MENU: False}), out)
    assert got == {"costos": [], "general": [], "consolidado": [],
                   "devoluciones": []}


# ------------------------------------------------------------------ 账单

def test_billing_resume_page_without_buttons(world):
    diag, out = world
    assert mf.download_billing_reports(Page(), out, months=2, diag=diag) == {}
    assert causes(diag, ["billing/1", "billing/2"]) == {
        "billing/1": "UI_NOT_FOUND", "billing/2": "UI_NOT_FOUND"}


def test_billing_fewer_periods_than_asked_and_one_that_breaks(world, monkeypatch):
    diag, out = world
    page = Page({mf.JS_CARD_INFO: {"month": "Agosto", "enCurso": False}},
                buttons=[El()])
    # 唯一的一期：Reportes 标签页找不到（deep_find 返回 None）
    got = mf.download_billing_reports(page, out, months=2, diag=diag)
    assert len(got) == 1 and list(got.values()) == [[]]   # 这一期留在结果里
    assert causes(diag, ["billing/1", "billing/2"]) == {
        "billing/1": "UI_NOT_FOUND", "billing/2": "NOT_AVAILABLE"}
    assert "Reportes tab not found" in diag.items["billing/1"]["evidence"]["exception"]
    # 老调用方不传 diag：异常照旧往外抛，交给 _guard
    with pytest.raises(RuntimeError):
        mf.download_billing_reports(page, out, months=1)


def test_billing_month_route_records_the_target_month_failing(world):
    diag, out = world
    shown = iter(["Septiembre", None])          # 目标月被重定向；下一期页面没有月份
    page = Page()
    page.answers = {k: None for k in ()}
    page.execute_script = lambda js, *a: (
        {"lang": "es-MX", "body": "", "alerts": []} if js == dd.JS_PAGE_FACTS
        else next(shown))
    got = mf.download_billing_for_months(page, out, 2020, 7, extra=2, diag=diag)
    assert got == {}
    expected = ["billing/2020-07", "billing/2020-08", "billing/2020-09"]
    recs = {r["form"]: r for r in dd.diagnose_store(expected, diag)}
    assert recs["billing/2020-07"]["cause"] == "PAGE_BLOCKED"      # 目标月：要人看
    assert "Septiembre" in recs["billing/2020-07"]["explain"]
    assert recs["billing/2020-08"]["cause"] == "PAGE_BLOCKED"
    # 第三期没去试，原因跟着第二期走
    assert recs["billing/2020-09"]["cause"] == "PAGE_BLOCKED"
    assert "billing/2020-08" in recs["billing/2020-09"]["explain"]


def test_billing_month_route_later_period_redirected_is_not_available(world):
    diag, out = world
    page = Page()
    page.execute_script = lambda js, *a: (
        {"lang": "es-MX", "body": "", "alerts": []} if js == dd.JS_PAGE_FACTS
        else "Julio")
    mf.download_billing_for_period(page, out, 2020, 8, diag=diag, required=False)
    assert causes(diag, ["billing/2020-08"]) == {"billing/2020-08": "NOT_AVAILABLE"}


# ------------------------------------------------------------------ Ventas

def test_ventas_button_disabled_is_recorded_with_candidates(world):
    diag, _out = world
    answers = {mf.JS_CURRENT_PERIOD: 2, mf.JS_WIDGET_ROWS: []}
    page = Page(answers)
    real = page.execute_script

    def script(js, *a):
        if "filter(e => /Descargar Excel de ventas/" in js:
            return [{"visible": True, "disabled": True}]
        return real(js, *a)
    page.execute_script = script
    assert mf.request_sales_excel(page, period_months=6, diag=diag) is None
    ev = diag.items["ventas"]["evidence"]
    assert ev["signal"] == "ui-disabled"
    assert ev["candidates"] == [{"visible": True, "disabled": True}]
    assert diag.flags["range_clamped"]["ventas"] == {"requested_months": 6,
                                                     "actual_months": 2}
    assert causes(diag, ["ventas"]) == {"ventas": "UI_NOT_FOUND"}


def test_ventas_probe_and_fetch_leave_what_they_saw_on_the_entry(world):
    _diag, out = world
    rows = [{"title": "t", "sub": "Generando", "date": "hoy", "hasDl": False}]
    entry = {"kind": "ventas", "rows_before": 0}
    page = Page({mf.JS_WIDGET_ROWS: rows}, url="https://x/ventas/omni/listado")
    assert mf.probe_sales_excel(page, entry) is False          # 返回值没变
    assert entry["_probe"] == {"row_state": "ventas-widget", "widget_rows": 1,
                               "rows_before": 0, "top_sub": "Generando",
                               "top_has_dl": False}
    assert mf.fetch_sales_excel(page, entry, out) == []
    assert entry["_fetch"] == "no-click-target"


# ------------------------------------------------------------------ MercadoPago

def test_mp_request_without_create_button(world):
    diag, _out = world
    page = Page({mp.JS_OPEN_CREATE_MENU: False},
                url="https://www.mercadopago.com.mx/balance/reports/collection")
    mp.LANG_CHECKED["done"] = True
    assert mp.request_report(page, report="collection", diag=diag) is None
    assert causes(diag, ["mercadopago/collection"]) == {
        "mercadopago/collection": "UI_NOT_FOUND"}
    assert "Crear reporte" in diag.items["mercadopago/collection"]["evidence"]["step"]
    # 不传 diag 也一样返回 None
    assert mp.request_report(page, report="collection") is None


def test_mp_request_adopting_the_top_row_is_flagged_not_failed(world, monkeypatch):
    _diag, _out = world
    day = datetime.date(2026, 10, 13)
    monkeypatch.setattr(mp, "_select_range",
                        lambda d, start, end: (day - datetime.timedelta(30), day))
    page = Page({mp.JS_OPEN_CREATE_MENU: True, mp.JS_CLICK_CREATE_ITEM: True,
                 mp.JS_GENERAR_STATE: {"disabled": False},
                 mp.JS_FIND_ROW: None,
                 mp.JS_AC_TOP_PERIOD: "12/sep/2026 a 12/oct/2026"})
    mp.LANG_CHECKED["done"] = True
    got = mp.request_report(page, report="collection")
    # 流程不变：照旧返回待取条目，按站点顶行的期间匹配
    assert got["range"] == "12/sep/2026 a 12/oct/2026" and got["kind"] == "mercadopago"
    assert got["adopted_existing_row"]["requested"] == "13/sep/2026 a 13/oct/2026"
    assert got["adopted_existing_row"]["adopted"] == "12/sep/2026 a 12/oct/2026"
    # 站点原样回显了所选区间：没有标记
    page.answers[mp.JS_FIND_ROW] = {"cells": [], "ready": False}
    assert "adopted_existing_row" not in mp.request_report(page, report="collection")


def test_poscobro_no_new_row_is_flagged_not_failed(world, monkeypatch):
    diag, _out = world
    day = datetime.date(2026, 10, 13)
    monkeypatch.setattr(mp, "_select_range",
                        lambda d, start, end: (day - datetime.timedelta(60), day))
    top = "14 agosto 2026 - 14 octubre 2026"
    page = Page({mp.JS_OPEN_CREATE_MENU: True, mp.JS_CLICK_CREATE_ITEM: True,
                 mp.JS_AC_CHECK_ONE: "clicked",
                 mp.JS_AC_OP_STATES: [{"t": "Reclamos", "chk": True}],
                 mp.JS_AC_CLICK_TEXT: True, mp.JS_AC_DD_TEXTS: ["Todos los estados"],
                 mp.JS_AC_TOP_PERIOD: top})
    mp.LANG_CHECKED["done"] = True
    got = mp.request_report(page, report="after_collection", diag=diag)
    assert got["range"] == top and got["flavor"] == "wizard"
    assert got["adopted_existing_row"] == {
        "requested": "2026-08-14 .. 2026-10-13", "adopted": top,
        "reason": "点 Generar 后列表顶行没有变化"}
    assert not diag.items                               # 不是失败
    # 向导第一步过不去：记原因，返回 None
    page.answers[mp.JS_AC_CLICK_TEXT] = lambda text: text != "siguiente"
    assert mp.request_report(page, report="after_collection", diag=diag) is None
    assert causes(diag, ["mercadopago/after_collection"]) == {
        "mercadopago/after_collection": "UI_NOT_FOUND"}


def test_account_statement_modal_never_opens(world):
    diag, _out = world
    mp.LANG_CHECKED["done"] = True
    assert mp.request_report(Page(), report="account_statement", diag=diag) is None
    assert causes(diag, ["mercadopago/account_statement"]) == {
        "mercadopago/account_statement": "UI_NOT_FOUND"}


ROWS = [
    (None, mp.WAITING, "absent"),
    ({"cells": ["a", "r", "c", "En preparación"], "ready": False, "state": "",
      "formats": [], "status": "En preparación"}, mp.WAITING, "generating"),
    ({"cells": ["a", "r"], "ready": True, "formats": [".csv"], "status": "x"},
     mp.READY, "ready"),
    ({"cells": ["a", "r"], "ready": False, "formats": [".xlsx"], "status": "x"},
     mp.READY, "ready"),
    ({"cells": ["a", "r"], "ready": False, "state": "empty", "formats": []},
     mp.EMPTY, "empty"),
    ({"cells": ["a", "r"], "ready": False, "state": "delayed", "formats": []},
     mp.DELAYED, "delayed"),
]


@pytest.mark.parametrize("row,answer,state", ROWS, ids=[r[2] + str(i) for i, r
                                                       in enumerate(ROWS)])
def test_mp_probe_answer_is_unchanged_and_the_row_is_recorded(world, row, answer,
                                                              state):
    entry = {"kind": "mercadopago", "report": "release", "url": "u", "range": "r",
             "label": "Liberaciones", "fmt": "csv"}
    assert mp.probe(Page({mp.JS_FIND_ROW: row}), entry) == answer
    assert entry["_probe"]["row_state"] == state


def test_mp_accordion_probe_records_the_action(world):
    entry = {"kind": "mercadopago", "report": "account_statement", "url": "u",
             "range": "Septiembre de 2026", "flavor": "accordion", "fmt": "csv"}
    page = Page({mp.JS_ACCT_FORMAT_STATE: {"head": "1 septiembre 2026 - 30",
                                           "action": "generar"}})
    assert mp.probe(page, entry) == mp.WAITING
    assert entry["_probe"] == {"row_state": "accordion", "action": "generar",
                               "head": "1 septiembre 2026 - 30"}
    page.answers[mp.JS_ACCT_FORMAT_STATE] = None
    assert mp.probe(page, entry) == mp.WAITING
    assert entry["_probe"]["action"] is None


def test_mp_fetch_says_whether_there_was_anything_to_click(world, monkeypatch):
    _diag, out = world
    monkeypatch.setattr(mf, "wait_for_downloads", lambda *a, **k: [])
    entry = {"kind": "mercadopago", "report": "release", "range": "r", "fmt": "csv",
             "label": "Liberaciones"}
    assert mp.fetch(Page({mp.JS_CLICK_ROW: None}), entry, out) == []
    assert entry["_fetch"] == "no-click-target"
    assert mp.fetch(Page({mp.JS_CLICK_ROW: ["a", "r"]}), entry, out) == []
    assert entry["_fetch"] == "clicked"
