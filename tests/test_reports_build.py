# -*- coding: utf-8 -*-
"""report_build.py 跑在一个最小的合成下载目录上：退出码、--strict、校验只打印一次。

不碰真实数据（仓库里也没有）。目录里只有报表必需的两个文件 —— Ventas 和一期
账单，各两笔订单，数字凑成所有 error 级校验都能过。
"""
import json
import os

import pytest
from openpyxl import Workbook

import report_build

NOTICE = ("No pudimos mostrar algunos datos. "
          "Intenta generando el Excel de ventas nuevamente.")

VENTAS_HEAD = [
    report_build.V_ORDER_ID, report_build.V_DATE, report_build.V_STATUS,
    report_build.V_UNITS, report_build.V_REVENUE, report_build.V_COMMISSION,
    report_build.V_SHIP_INCOME, report_build.V_SHIP_COST, report_build.V_REFUND,
    report_build.V_TOTAL, report_build.V_SKU, report_build.V_TITLE,
]
# 佣金 15% + 代扣税 9%（校验要求税率落在 8.5%–9.5%），运费 10%。
VENTAS_ROWS = [
    ("2000000000000001", "3 de agosto de 2026 09:24 hs.", "Entregado", 1,
     100, -24, 0, -10, 0, 66, "SKU-A", "Producto A"),
    ("2000000000000002", "5 de agosto de 2026 10:00 hs.", "Entregado", 2,
     200, -48, 0, -20, 0, 132, "SKU-B", "Producto B"),
]
BILLING_HEAD = [
    report_build.F_INVOICE, report_build.F_CHARGE_DATE, report_build.F_CHARGE_NO,
    report_build.F_DETALLE, report_build.F_CHANNEL, report_build.F_STATE,
    report_build.F_REVERSES, report_build.F_AMOUNT, report_build.F_SALE_NO,
    report_build.F_PACK_NO, report_build.F_SALE_DATE,
]
BILLING_ROWS = [
    ("F1", "2026-08-03", "9001", "Cargo por venta", "Si", "Facturado", None,
     15, "2000000000000001", None, "2026-08-03"),
    ("F1", "2026-08-03", "9002", "Cargo por envíos de Mercado Libre", "Si",
     "Facturado", None, 10, "2000000000000001", None, "2026-08-03"),
    ("F1", "2026-08-05", "9003", "Cargo por venta", "Si", "Facturado", None,
     30, "2000000000000002", None, "2026-08-05"),
    ("F1", "2026-08-05", "9004", "Cargo por envíos de Mercado Libre", "Si",
     "Facturado", None, 20, "2000000000000002", None, "2026-08-05"),
]


def make_download_folder(folder, banner=False, commission=(-24, -48)):
    """写一个够出报表的下载目录。banner=True 时在 Ventas 顶部插入平台的横幅。

    横幅占一行，把分组行和表头各往下推一行 —— 和生产上 9/15 TOOL_TA03 那份
    导出的形状一致。commission 可以改掉，用来造一条不过的校验（税率越界）。
    """
    folder = str(folder)
    if not os.path.isdir(folder):
        os.makedirs(folder)

    wb = Workbook()
    ws = wb.active
    ws.title = report_build.VENTAS_SHEET
    ws["A1"] = "Ventas"
    shift = 0
    if banner:
        ws["A4"] = NOTICE
        shift = 1
    group = report_build.VENTAS_HEADER_ROW + shift        # 表头的上一行（1-indexed）
    ws.cell(group, 1, "Ventas")
    ws.cell(group, 11, "Publicaciones")
    for col, value in enumerate(VENTAS_HEAD, 1):
        ws.cell(group + 1, col, value)
    for r, (row, com) in enumerate(zip(VENTAS_ROWS, commission), group + 2):
        row = list(row)
        row[9] += com - row[5]
        row[5] = com
        for col, value in enumerate(row, 1):
            ws.cell(r, col, value)
    wb.save(os.path.join(folder, "20260901_Ventas_MX_Mercado_Libre_y_Mercado_Shops_"
                                 "2026-09-01_10-00hs_123456789.xlsx"))

    wb = Workbook()
    ws = wb.active
    ws.title = "REPORT"
    ws["A1"] = "Reporte de facturación"
    head = report_build.BILLING_HEADER_ROW + 1
    for col, value in enumerate(BILLING_HEAD, 1):
        ws.cell(head, col, value)
    for r, row in enumerate(BILLING_ROWS, head + 1):
        for col, value in enumerate(row, 1):
            ws.cell(r, col, value)
    wb.save(os.path.join(folder, "Reporte_Facturacion_MercadoLibre_Ago2026.xlsx"))
    return folder


def run_build(folder, out, *extra):
    return report_build.run(["--month", "2026-08", "--store", "TEST=%s" % folder,
                             "--out", str(out), "--no-recalc"] + list(extra))


def checks_from(output):
    line = [l for l in output.splitlines() if l.startswith("[REPORT_CHECKS_JSON] ")]
    assert len(line) == 1
    return json.loads(line[0].split(" ", 1)[1])


def test_clean_folder_writes_report_and_exits_zero(tmp_path, capsys):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    out = tmp_path / "out" / "r.xlsx"
    os.makedirs(str(tmp_path / "out"))

    assert run_build(folder, out) == report_build.EXIT_OK
    assert out.is_file()
    checks = checks_from(capsys.readouterr().out)
    assert not [c for c in checks if not c["ok"] and not c["skipped"]]


def test_failed_check_is_printed_once_with_its_hint(tmp_path, capsys):
    # 税只扣了 5%：tax_rate 校验不过。文件照常写出，退出码 1。
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000",
                                  commission=(-20, -40))
    out = tmp_path / "r.xlsx"

    assert run_build(folder, out) == report_build.EXIT_CHECKS
    assert out.is_file()
    output = capsys.readouterr().out
    failed = [c for c in checks_from(output) if not c["ok"] and not c["skipped"]]
    assert [c["key"] for c in failed] == ["tax_rate"]
    human = "\n".join(l for l in output.splitlines()
                      if not l.startswith("[REPORT_CHECKS_JSON] "))
    # 说明文字只出现一次（以前 [校验] 和 [诊断] 各一次），提示紧跟在它下面。
    assert human.count(failed[0]["detail"]) == 1
    assert human.count("排查提示") == 1
    assert "[诊断]" not in human
    lines = human.splitlines()
    at = next(i for i, l in enumerate(lines) if failed[0]["detail"] in l)
    assert "排查提示" in lines[at + 1]


def test_banner_is_a_warning_without_strict(tmp_path, capsys):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000", banner=True)
    out = tmp_path / "r.xlsx"

    # 不 strict 时行为不变：照常出报表，warn 级不影响退出码。
    assert run_build(folder, out) == report_build.EXIT_OK
    assert out.is_file()
    failed = [c for c in checks_from(capsys.readouterr().out) if not c["ok"]]
    assert [(c["key"], c["severity"]) for c in failed] == [("ventas_incomplete", "warn")]


def test_strict_blocks_a_platform_declared_incomplete_export(tmp_path, capsys):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000", banner=True)
    out = tmp_path / "r.xlsx"

    assert run_build(folder, out, "--strict") == report_build.EXIT_STRICT_ABORT
    assert not out.exists()
    assert "ventas_incomplete" in capsys.readouterr().out


def test_strict_still_writes_a_clean_report(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")
    out = tmp_path / "r.xlsx"

    assert run_build(folder, out, "--strict") == report_build.EXIT_OK
    assert out.is_file()


def test_strict_blocks_error_level_checks(tmp_path):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000",
                                  commission=(-20, -40))
    out = tmp_path / "r.xlsx"

    assert run_build(folder, out, "--strict") == report_build.EXIT_STRICT_ABORT
    assert not out.exists()


def test_unhandled_exception_gets_its_own_exit_code(tmp_path, capsys, monkeypatch):
    folder = make_download_folder(tmp_path / "TEST" / "20260901_100000")

    def boom(*args, **kwargs):
        raise KeyError("columna que no existe")

    monkeypatch.setattr(report_build, "build_workbook", boom)
    out = tmp_path / "r.xlsx"

    code = run_build(folder, out)
    assert code == report_build.EXIT_CRASH
    assert code not in (report_build.EXIT_OK, report_build.EXIT_CHECKS)
    output = capsys.readouterr().out
    assert "Traceback (most recent call last)" in output
    assert "columna que no existe" in output


def test_no_usable_store_keeps_its_exit_code(tmp_path, capsys):
    folder = str(tmp_path / "TEST" / "20260901_100000")
    make_download_folder(folder)
    # 表头整行抹掉列名：没有销售日期，这家店被剔除，没有店可用。
    from openpyxl import load_workbook
    path = [os.path.join(folder, f) for f in os.listdir(folder) if "Ventas" in f][0]
    wb = load_workbook(path)
    ws = wb[report_build.VENTAS_SHEET]
    ws.cell(report_build.VENTAS_HEADER_ROW + 1, 2, "Sale date")
    wb.save(path)

    assert run_build(folder, tmp_path / "r.xlsx") == report_build.EXIT_NO_STORE
    assert not (tmp_path / "r.xlsx").exists()


def test_combined_report_survives_two_stores_with_equal_rates(tmp_path):
    # 两家店扣费占比完全相同（生产上最常见的是两家当月都没有销售，都是 0）。
    # trend_notes 以前拿 (比率, 店铺 dict) 直接排序，比率相等就去比 dict，
    # TypeError 把整份汇总报表带走。
    a = make_download_folder(tmp_path / "A" / "20260901_100000")
    b = make_download_folder(tmp_path / "B" / "20260901_100000")
    out = tmp_path / "all.xlsx"

    code = report_build.run(["--month", "2026-08", "--store", "A=%s" % a,
                             "--store", "B=%s" % b, "--out", str(out), "--no-recalc"])
    assert code == report_build.EXIT_OK
    assert out.is_file()
