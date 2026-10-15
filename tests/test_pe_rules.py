# -*- coding: utf-8 -*-
"""秘鲁规则的单元测试：PEN 格式与容差、IGV、自然月账单、"不代扣"、没有 Full、
秘鲁专名的替换，以及入口脚本认得秘鲁。全部用合成数据。
"""
import json
import os

import pandas as pd
import pytest

import financial_report
import report_build as rb
import run_reports


@pytest.fixture(autouse=True)
def _back_to_mx():
    yield
    rb.set_site("mx")


T = pd.Timestamp


# ---------------------------------------------------------------- 站点常量

def test_set_site_pe_columns_money_and_tolerances():
    assert rb.set_site("PE") == "pe"
    assert rb.VENTAS_SHEET == "Ventas PE" and rb.VENTAS_GLOB == "*Ventas_PE*.xlsx"
    assert rb.V_REVENUE == "Ingresos por productos (PEN)"
    assert rb.V_TOTAL == "Total (PEN)"
    assert (rb.CUR, rb.CUR_CN, rb.COUNTRY_CN) == ("PEN", "秘鲁索尔", "秘鲁")
    # 索尔有分：两位小数，容差按分
    assert rb.MNY == "#,##0.00;(#,##0.00);-"
    assert rb.fmt(1234567.456) == "1,234,567.46"
    assert (rb.TOL_CENT, rb.CN_TOL, rb.BRIDGE_A_TOL) == (0.01, 0.05, 0.05)
    assert rb.IVA_RATE == 0.18 and rb.VAT_NAME == "IGV" and rb.TAX_AUTH == "SUNAT"
    assert rb.WITHHOLDING is False
    assert rb.HAS_FULL is False and rb.BONIF_COL is True


def test_switching_back_to_mx_or_cl_restores_their_flags():
    """秘鲁新增的键（HAS_FULL 等）不能"漏"到下一份墨西哥/智利报表里。"""
    rb.set_site("pe")
    rb.set_site("mx")
    assert (rb.HAS_FULL, rb.BONIF_COL, rb.VAT_NAME, rb.IVA_RATE) == (True, False, "IVA", 0.16)
    assert rb.WITHHOLDING is True and rb.BILL_CYCLE_DAY == 1
    rb.set_site("pe")
    rb.set_site("cl")
    assert (rb.HAS_FULL, rb.BONIF_COL, rb.VAT_NAME, rb.TAX_AUTH) == (True, False, "IVA", "SII")


def test_pe_bill_is_the_calendar_month():
    rb.set_site("pe")
    assert rb.BILL_CYCLE_DAY == 1
    assert rb.bill_cycles(2026, 9) == [(2026, 9, T(2026, 9, 1), T(2026, 9, 30))]


def test_setiembre_file_and_date_tokens():
    # 实测文件名仍是 Sept；平台若按秘鲁拼法给 Set 也要认得
    assert rb.file_month_token("Reporte_Facturacion_MercadoLibre_Sept2026.xlsx") == (2026, 9)
    assert rb.file_month_token("Reporte_Facturacion_MercadoLibre_Set2026.xlsx") == (2026, 9)
    assert rb.parse_spanish_date("30 de setiembre de 2026 10:00 hs.") == T(2026, 9, 30, 10, 0)
    assert financial_report.FILE_MONTH_TOKENS["set"] == 9


# ---------------------------------------------------------------- 不代扣 / IGV

def _S(tax_ret, dev, ing=27859.94, mp_tax=0.0):
    return {"store": "X", "mp": {"tax": mp_tax} if mp_tax is not None else None,
            "m": {"tax_ret": tax_ret, "vcol_dev": dev, "vcol_rest": tax_ret - dev,
                  "ing": ing, "com_match_n": 202, "com_match_of": 202}}


def test_pe_no_withholding_check_and_localized_wording():
    rb.set_site("pe")
    ok = rb._no_withholding_check(_S(0.0, 0.0))
    assert ok["ok"] and ok["severity"] == "error"
    # 若平台开始代扣（墨西哥式 ~9%）或流水出现税金：失败
    assert not rb._no_withholding_check(_S(0.09 * 27859.94, 0.0))["ok"]
    assert not rb._no_withholding_check(_S(0.0, 0.0, mp_tax=-12.5))["ok"]
    # 文字在 validate() 出口统一换成秘鲁的专名
    assert rb.localize(ok["name"]) == "平台未代扣税（秘鲁）"
    assert "IGV" in rb.localize(ok["detail"]) and "IVA" not in rb.localize(ok["detail"])


def test_localize_only_touches_peru():
    s = "智利平台不代扣 IVA，向 SII 申报（× 19/119），电子发票（DTE）；CLP 比索；ACTIVA 不动"
    rb.set_site("mx")
    assert rb.localize(s) == s
    rb.set_site("cl")
    assert rb.localize(s) == s
    rb.set_site("pe")
    got = rb.localize(s)
    assert got.startswith("秘鲁平台不代扣 IGV，向 SUNAT 申报（× 18/118）")
    assert "CPE" in got and "索尔" in got and "ACTIVA" in got
    assert rb.localize(None) is None and rb.localize(3.5) == 3.5


def test_pe_pl_rows_same_structure_as_mx_with_flex_note():
    rb.set_site("mx")
    mx_keys = [r[0] for r in rb.build_pl_rows()]
    rb.set_site("pe")
    rows = rb.build_pl_rows()
    assert [r[0] for r in rows] == mx_keys                    # 行结构不变
    env = [r for r in rows if r[0] == "env_ing"][0]
    assert "Descuentos y bonificaciones" in env[2] and "Flex" in env[4]
    notes = dict((r[0], r[4]) for r in rows)
    assert "不代扣" in notes["com_neg"]


def test_pe_memo_says_calendar_month_not_two_bills():
    rb.set_site("pe")
    memo = rb.pl_memo_no_withholding([{"store": "X", "cycle_day": 1,
                                       "m": {"vcol_dev": 0.0, "vcol_rest": 0.0}}])
    pay = [x for x in memo if x[1].startswith("Pago con")][0]
    assert "两张账单" not in pay[4] and "自然月" in pay[4]
    assert any("IVA" in x[0] for x in memo)        # 写出工作簿时由 localize 换成 IGV


# ---------------------------------------------------------------- 入口脚本

def test_financial_report_pe_needs_one_bill_and_accepts_ventas_pe(tmp_path):
    with open(os.path.join(str(tmp_path), "run_meta.json"), "w") as fh:
        json.dump({"site": "pe"}, fh)
    assert financial_report.bills_needed(str(tmp_path), (2026, 12)) == {(2026, 12)}
    for n in ("20261015_Ventas_PE_x.xlsx", "Reporte_Facturacion_MercadoLibre_Sept2026.xlsx"):
        open(os.path.join(str(tmp_path), n), "w").close()
    assert financial_report._usable(str(tmp_path)) == (True, [])
    assert financial_report.billing_months(str(tmp_path)) == {(2026, 9)}


def test_run_reports_names_the_peru_workbook():
    assert run_reports.combined_name("2026-09", "PE") == "PE_ML_全店汇总_2026-09.xlsx"
    assert run_reports.combined_name("2026-09", "pe") == "PE_ML_全店汇总_2026-09.xlsx"
    assert run_reports.combined_name("2026-09") == "MX_ML_全店汇总_2026-09.xlsx"


def test_run_reports_accepts_region_pe(monkeypatch, capsys):
    """--region PE 能过参数解析，并把 pe 当站点传给 report_build。"""
    seen = {}
    monkeypatch.setattr(run_reports, "store_names",
                        lambda stores, region: seen.setdefault("region", region) and [])
    monkeypatch.setattr(run_reports.store_config, "region_name", lambda r: "秘鲁")
    monkeypatch.setattr("sys.argv", ["run_reports.py", "--region", "pe", "--month", "2026-09"])
    with pytest.raises(run_reports.StoreConfigError):
        run_reports.main()
    assert seen["region"] == "PE"
