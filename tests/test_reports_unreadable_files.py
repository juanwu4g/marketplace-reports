# -*- coding: utf-8 -*-
"""订单基准：可选文件读不出来（平台把出错网页存成了 .xlsx/.csv）不能让报表炸掉。

生产实况（2026-10-30，EWTTO_SM 2026-09）：五份库存类文件都是网页，pandas 报
"Excel file format cannot be determined"。账单基准照常出表、⑩ 页记一条
load:<文件> 校验、评级有缺口；订单基准在 load_storage 炸掉，这家店的订单基准
报表和整个墨西哥的订单基准全店汇总都没写出来。本文件钉住订单基准的新行为：

  * 读不出来的可选文件当它不存在，文件名和错误进 S["load_errors"]；
  * validate() 为每一条记 load:<文件名> 校验（名称、措辞与账单基准相同）；
    库存类文件和 Cargos Full 是 warn 级（② 的金额不用它们），Pagos 和
    MercadoPago 流水是 error 级；
  * warn 级的 load: 校验不把评级拉成有缺口；工作簿照常写出，⑩ 页能看到它；
  * Ventas / Facturación 读不出来照旧抛异常 —— 没有它们报表就不存在。

全部合成数据，不碰真实文件。
"""
import os

import openpyxl
import pytest

from core import grading, results
from platforms.mercadolibre import report_build
from platforms.mercadolibre.report import loading
from platforms.mercadolibre.report.catalog import SHEET_NAMES
from test_reports_build import checks_from, make_download_folder, run_build

# 平台出错页的样子：几行 HTML，没有一个分号，不是 zip —— openpyxl/pandas 都认不出来。
HTML = (b"<!DOCTYPE html>\n<html><head><title>Error</title></head>\n"
        b"<body><h1>500 Internal Server Error</h1>\n<p>Intenta de nuevo.</p></body></html>\n")

# 生产那天的五个文件名模式（会计月改成夹具用的 2026-08）。
STORAGE_IN_MONTH = "01-08-26_31-08-26_Costos_por_servicio_almacenamiento.xlsx"
STORAGE_NEXT_MONTH = "01-09-26_30-09-26_Costos_por_servicio_almacenamiento.xlsx"
STOCK = "stock_general_full_123456789_0a1b2c3d.xlsx"
CONCILIATION = "conciliation_SELLER_123456789_0a1b2c3d.xlsx"
RETURNS = "Returns_30-10-2026.xlsx"
STOCK_TYPE_FILES = (STORAGE_IN_MONTH, STORAGE_NEXT_MONTH, STOCK, CONCILIATION, RETURNS)

# 另外几份可选文件：账单组的两份和 MercadoPago 流水。
PAGOS = "Reporte_Pagos_Facturas_Ago2026.xlsx"
CARGOS = "Reporte_Cargos_Full_Ago2026.xlsx"
SETTLEMENT = "settlement_v2-SELLER-manual-2026-09-01-100000.csv"


def _garbage(folder, names):
    for n in names:
        with open(os.path.join(str(folder), n), "wb") as f:
            f.write(HTML)


def _derive(folder):
    """真走 derive -> validate，不拿手写的 check 字典冒充。"""
    p0, p1 = report_build.month_bounds("2026-08")
    S = report_build.derive("TEST", str(folder), p0, p1)
    return S, report_build.validate([S])


def _load_checks(checks):
    return [c for c in checks if c["key"].startswith("load:")]


def _grade(checks):
    return grading.grade("ML", "order", True, checks)


# ── 库存类文件（生产那天的五份）───────────────────────────────────────

def test_stock_type_garbage_does_not_crash_and_is_recorded(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, STOCK_TYPE_FILES)

    S, checks = _derive(folder)             # 以前在 load_storage 的 pd.read_excel 炸掉

    names = [n for n, _ in S["load_errors"]]
    # 墨西哥的 load_storage 只打开与会计月重叠的那份；9 月那份根本没碰。
    # conciliation 订单基准从不读（只有账单基准的 ⑨ 用它），所以也记不到。
    assert names == [STORAGE_IN_MONTH, RETURNS, STOCK]
    for _, err in S["load_errors"]:
        assert err                           # 原始错误原样带上，排查用
    # 当它们不存在：文件槽位空、数据为空，报表其余部分照常算
    assert S["files"]["storage"] is None and S["files"]["stock"] is None
    assert S["files"]["returns"] is None
    assert len(S["stock"]) == 0 and len(S["returns"]) == 0 and len(S["returns_all"]) == 0
    assert S["m"]["storage_accrual"] == 0.0 and S["m"]["storage_range"] is None


def test_load_checks_are_warn_level_in_bill_basis_wording(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, STOCK_TYPE_FILES)

    S, checks = _derive(folder)

    hit = _load_checks(checks)
    assert [c["key"] for c in hit] == ["load:" + n for n, _ in S["load_errors"]]
    for c, (name, err) in zip(hit, S["load_errors"]):
        assert c["ok"] is False and c["skipped"] is False
        assert c["severity"] == "warn"
        assert c["store"] == "TEST"
        assert c["name"] == "读取 %s" % name                 # 与账单基准同一措辞
        assert c["detail"].startswith("读不出来：%s" % err)
    # 排在这家店的校验最前面：读不出来是后面"跳过/没有文件"的原因
    assert checks[0]["key"].startswith("load:")
    # 没有任何 error 级未通过 —— 库存类文件不碰 ② 的金额
    assert not [c for c in checks if not c["ok"] and not c["skipped"] and c["severity"] == "error"]
    # 排查提示也有
    assert report_build.check_hints(hit[0], S)


def test_warn_level_load_checks_do_not_make_the_grade_gaps(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, STOCK_TYPE_FILES)

    S, checks = _derive(folder)

    g = _grade(checks)
    assert g["level"] == results.COMPLETE, g        # 评级只看 error 级校验
    assert g["level"] != results.GAPS


def test_garbage_stock_files_change_no_number(tmp_path):
    plain, _ = _derive(make_download_folder(tmp_path / "a" / "20260901_100000"))
    folder = make_download_folder(tmp_path / "b" / "20260901_100000")
    _garbage(folder, STOCK_TYPE_FILES)
    flagged, _ = _derive(folder)

    numbers = {k: v for k, v in plain["m"].items() if isinstance(v, (int, float))}
    assert numbers
    assert {k: v for k, v in flagged["m"].items() if isinstance(v, (int, float))} == numbers


def test_workbook_is_written_and_sheet_10_lists_the_files(tmp_path, capsys):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, STOCK_TYPE_FILES)
    out = tmp_path / "r.xlsx"

    assert run_build(folder, out) == report_build.EXIT_OK      # warn 级不算校验未过
    assert out.is_file()
    checks = checks_from(capsys.readouterr().out)
    hit = _load_checks(checks)
    assert {c["key"] for c in hit} == {"load:" + n for n in (STORAGE_IN_MONTH, RETURNS, STOCK)}
    assert all(c["severity"] == "warn" and not c["ok"] for c in hit)

    ws = openpyxl.load_workbook(str(out))[SHEET_NAMES[9]]
    text = "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None)
    for n in (STORAGE_IN_MONTH, RETURNS, STOCK):
        assert "读取 %s" % n in text
    assert "读不出来：" in text


def test_a_clean_folder_records_nothing(tmp_path):
    S, checks = _derive(make_download_folder(tmp_path / "TEST" / "20260901_100000"))
    assert S["load_errors"] == []
    assert not _load_checks(checks)


# ── 账单组的可选文件与 MercadoPago 流水 ─────────────────────────────────

def test_pagos_cargos_and_settlement_garbage(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, (PAGOS, CARGOS, SETTLEMENT))

    S, checks = _derive(folder)

    names = [n for n, _ in S["load_errors"]]
    assert sorted(names) == sorted([PAGOS, CARGOS, SETTLEMENT])
    assert names.count(SETTLEMENT) == 1            # 三个函数读同一份流水，只记一次
    assert S["files"]["pagos"] is None and S["files"]["cargos_full"] is None
    assert S["files"]["settlement"] is None and S["mp"] is None and S["mp_orders"] is None
    assert S["m"]["pay_auto"] == 0.0 and S["m"]["cargos_full_report_total"] == 0.0

    sev = {c["key"]: c["severity"] for c in _load_checks(checks)}
    assert sev["load:" + CARGOS] == "warn"          # 只用来交叉核对，② 不变
    assert sev["load:" + PAGOS] == "error"          # ② 的"本期账单实际支付"行、桥 B
    assert sev["load:" + SETTLEMENT] == "error"     # 订单基准用流水修正部分退款
    by_key = {c["key"]: c for c in checks}
    assert by_key["bridge_b"]["skipped"] and by_key["cargos_full"]["skipped"]
    assert _grade(checks)["level"] == results.GAPS


def test_a_settlement_csv_that_is_a_web_page_is_unreadable(tmp_path):
    # read_csv 能把网页"读"成一列垃圾而不报错 —— 没有流水的列就是读不出来。
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, (SETTLEMENT,))

    S, _ = _derive(folder)

    assert [n for n, _ in S["load_errors"]] == [SETTLEMENT]
    assert "TRANSACTION_DATE" in S["load_errors"][0][1]


# ── 老调用方：不传 errors 也照常（安静跳过）──────────────────────────────

def test_loaders_without_an_error_list_still_skip_quietly(tmp_path):
    _garbage(tmp_path, STOCK_TYPE_FILES + (PAGOS, CARGOS))
    p0, p1 = report_build.month_bounds("2026-08")
    assert loading.load_storage(str(tmp_path), p0, p1) == (None, 0.0, None)
    path, r = loading.load_returns(str(tmp_path))
    assert path is None and len(r) == 0
    path, snap, st = loading.load_stock(str(tmp_path))
    assert path is None and snap == "" and len(st) == 0
    assert loading.load_pagos(str(tmp_path), 2026, 8)[0] is None
    assert loading.load_cargos_full(str(tmp_path), 2026, 8) == (None, {}, 0.0)


# ── 定义报表的两份文件照旧报错 ──────────────────────────────────────────

def test_a_garbage_ventas_file_still_raises(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    ventas = [f for f in os.listdir(folder) if "Ventas" in f]
    assert len(ventas) == 1
    _garbage(folder, ventas)

    with pytest.raises(Exception):
        _derive(folder)


def test_a_garbage_billing_file_still_raises(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    _garbage(folder, ("Reporte_Facturacion_MercadoLibre_Ago2026.xlsx",))

    with pytest.raises(Exception):
        _derive(folder)
