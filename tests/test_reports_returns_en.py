# -*- coding: utf-8 -*-
"""Returns（Triages）导出现在是英文的：读进来要翻回西语，退货分析才不会是空的。"""
import openpyxl

from platforms.mercadolibre import report_build as rb


def _english_returns(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Triages"
    ws.append(["Report of returns in the warehouse"])
    ws.append(["How to read this report"])
    ws.append(["Order", "SKU", "IMEI", "Review date", "Product status",
               "Review result", "Money status"])
    ws.append([])
    ws.append(["2000017268739246", "XAMH76322", "-", "15-07-2026",
               "It does not work correctly",
               "Product to be picked up at distribution center", "We gave you the money"])
    ws.append(["2000017143809420", "JHDW40231", "-", "16-07-2026", "Other attribute",
               "Product is once again available for sale", "We refunded the buyer"])
    wb.save(path)


def test_english_returns_are_read_as_spanish(tmp_path):
    _english_returns(str(tmp_path / "Returns_13-10-2026.xlsx"))
    _, r = rb.load_returns(str(tmp_path))
    assert len(r) == 2 and r["fr"].notna().all(), "质检日要能解析"
    assert list(r["Estado del dinero"]) == [rb.MONEY_KEPT, rb.MONEY_REFUNDED]
    assert r["Resultado de la revisión"].iloc[0] == "Producto para retirar en centro de distribución"
    assert r["Estado del producto"].iloc[0] == "No funciona correctamente"
    assert r["Estado del producto"].iloc[1] == "Other attribute", "没见过的取值原样保留"
    assert "Número de orden" in r.columns
