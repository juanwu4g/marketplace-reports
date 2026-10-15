# -*- coding: utf-8 -*-
"""秘鲁报表跑在一个合成下载目录上：10 页齐全、各座桥平、校验按秘鲁的事实过 / 跳过。

全部是合成数据（真实下载含买家姓名、地址、证件号，不进仓库）。目录的形状照
IDG_PE 2026-09 的真实下载：
  * Ventas 页叫 "Ventas PE"，金额列带 (PEN)，两位小数；
  * "Descuentos y bonificaciones" 列**有值** —— Flex 配送补贴，MercadoPago 里是一笔
    单独的 CASHBACK；买家另付的运费是单独一笔 SETTLEMENT_SHIPPING；
  * 佣金 = 售价 × 类目费率（12%），Ventas 佣金列与账单 Cargo por venta 逐单相等，
    MercadoPago TAXES_AMOUNT 全为 0 —— 平台不代扣；
  * 账单按自然月（Sept2026 文件的计费日都在 9 月），发票号是 B001-…（boleta）；
  * 没有任何 Full / 仓储 / 库存 / Returns 文件 —— 秘鲁没有 Full。
"""
import contextlib
import io
import json
import os

import pytest
from openpyxl import Workbook, load_workbook

import report_build as rb


@pytest.fixture(autouse=True)
def _back_to_mx():
    yield
    rb.set_site("mx")


V_HEAD = ["# de venta", "Fecha de venta", "Estado", "Paquete de varios productos", "Unidades",
          "Ingresos por productos (PEN)", "Cargo por venta e impuestos (PEN)",
          "Ingresos por envío (PEN)", "Costos de envío (PEN)", "Descuentos y bonificaciones",
          "Anulaciones y reembolsos (PEN)", "Total (PEN)", "Orden de compra", "SKU",
          "Título de la publicación", "Forma de entrega"]
# (订单号, 日期, 售价, 佣金, 买家运费, 配送费, Flex 补贴, Total, SKU, 配送方式)
ORDERS = [
    ("2000000000000001", "2 de septiembre de 2026 10:00 hs.", 118, -14.16, None, None, 4.35,
     108.19, "SKU-A", "Mercado Envíos Flex"),
    ("2000000000000002", "10 de septiembre de 2026 12:30 hs.", 200, -24.0, 3.9, -7.7, None,
     172.2, "SKU-B", "Correo y puntos de despacho"),
    ("2000000000000003", "21 de septiembre de 2026 18:05 hs.", 100, -10.0, None, -5.0, None,
     85.0, "SKU-A", "Correo y puntos de despacho"),
]
B_HEAD = ["N° de factura fiscal", "Fecha del cargo", "Número del cargo", "Detalle",
          "Descontado de la operación", "Estado del cargo", "Cargo que bonifica",
          "Valor del cargo", "Número de venta", "Pago", "Fecha de venta", "Número de paquete"]
BILL_SEPT = [
    ("B001-00000001", "2026-09-02", "701", "Cargo por venta", "Si", None, None, 14.16,
     "2000000000000001", "P1", "2026-09-02", None),
    ("B001-00000001", "2026-09-10", "702", "Cargo por venta", "Si", None, None, 24.0,
     "2000000000000002", "P2", "2026-09-10", None),
    ("B001-00000001", "2026-09-10", "703", "Cargo por envíos de Mercado Libre", "Si", None, None,
     7.7, "2000000000000002", "P2", "2026-09-10", None),
    ("B001-00000001", "2026-09-21", "704", "Cargo por venta", "Si", None, None, 10.0,
     "2000000000000003", "P3", "2026-09-21", None),
    ("B001-00000001", "2026-09-21", "705", "Cargo por envíos de Mercado Libre", "Si", None, None,
     5.0, "2000000000000003", "P3", "2026-09-21", None),
    ("B001-00000001", "2026-09-01", "706", "Cargo por campaña de publicidad de Product Ads", "No",
     None, None, 50.0, None, None, None, None),
    ("B001-00000001", "2026-09-15", "707", "Cargo por campaña de publicidad de Product Ads", "No",
     None, None, 30.0, None, None, None, None),
    ("B001-00000001", "2026-09-01", "708", "Cargo por mantenimiento de Mi página", "No",
     "Anulado en factura", None, 39.0, None, None, None, None),
    ("B001-00000001", "2026-09-03", "709", "Anulación del cargo por mantenimiento de Mi página",
     "No aplica", None, "708", -39.0, None, None, None, None),
]
BILL_OCT = [
    ("Aún no disponible", "2026-10-01", "801", "Cargo por campaña de publicidad de Product Ads",
     "No", None, None, 20.0, None, None, None, None),
]
SETTLEMENT = [
    # TRANSACTION_DATE, SOURCE_ID, TYPE, AMOUNT, FEE, REAL, TAXES
    ("2026-09-02T10:00:05.000-05:00", "P1", "SETTLEMENT", 118.00, -14.16, 103.84, 0.0),
    ("2026-09-02T10:05:00.000-05:00", "C1", "CASHBACK", 4.35, 0.0, 4.35, 0.0),
    ("2026-09-10T12:30:05.000-05:00", "P2", "SETTLEMENT", 200.00, -31.70, 168.30, 0.0),
    ("2026-09-10T12:30:06.000-05:00", "S2", "SETTLEMENT_SHIPPING", 3.90, 0.0, 3.90, 0.0),
    ("2026-09-21T18:05:05.000-05:00", "P3", "SETTLEMENT", 100.00, -15.00, 85.00, 0.0),
]


def _billing_file(path, rows, banner=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "REPORT"
    ws["A1"] = "Fecha de actualización"
    if banner:
        ws["A3"] = "Este reporte podría tener cambios hasta el cierre del período de facturación"
    for i, h in enumerate(B_HEAD, 1):
        ws.cell(rb.BILLING_HEADER_ROW + 1, i, h)
    for r, row in enumerate(rows, rb.BILLING_HEADER_ROW + 2):
        for i, v in enumerate(row, 1):
            ws.cell(r, i, v)
    wb.save(path)


def make_pe_folder(folder, orders=ORDERS):
    folder = str(folder)
    os.makedirs(folder, exist_ok=True)
    with io.open(os.path.join(folder, "run_meta.json"), "w", encoding="utf-8") as fh:
        json.dump({"site": "pe", "month": "2026-09", "sales_window_months": 2,
                   "downloaded_at": "2026-10-15T13:30:22"}, fh)

    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas PE"
    ws.cell(rb.VENTAS_HEADER_ROW, 1, "Ventas")
    for i, h in enumerate(V_HEAD, 1):
        ws.cell(rb.VENTAS_HEADER_ROW + 1, i, h)
    for r, (oid, d, ing, com, envi, envc, bon, tot, sku, how) in enumerate(
            orders, rb.VENTAS_HEADER_ROW + 2):
        row = [oid, d, "Entregado", "No", 1, ing, com, envi, envc, bon, None, tot, None,
               sku, "Producto " + sku, how]
        for i, v in enumerate(row, 1):
            ws.cell(r, i, v)
    wb.save(os.path.join(folder, "20261015_Ventas_PE_Mercado_Libre_y_Mercado_Shops_"
                                 "2026-10-14_23-31hs_123456789.xlsx"))

    _billing_file(os.path.join(folder, "Reporte_Facturacion_MercadoLibre_Sept2026.xlsx"), BILL_SEPT)
    _billing_file(os.path.join(folder, "Reporte_Facturacion_MercadoLibre_Oct2026.xlsx"), BILL_OCT,
                  banner=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pagos y notas de crédito"
    head = ["Número de pago", "Número de nota de crédito", "Tipo de pago", "Medio de pago",
            "Fecha de pago / Emisión de NC", "Estado", "Fecha de anulación", "Importe total",
            "Importe aplicado a este mes"]
    for i, h in enumerate(head, 1):
        ws.cell(rb.PAGOS_HEADER_ROW + 1, i, h)
    for i, v in enumerate(["900001", None, "Pago con débito automático", "Mercado Pago",
                           "2026-10-06", "Pago aprobado", None, 80.0, 80.0], 1):
        ws.cell(rb.PAGOS_HEADER_ROW + 2, i, v)
    wb.create_sheet("Detalle de Pagos del mes")
    wb.save(os.path.join(folder, "Reporte_Pagos_Facturas_Sept2026.xlsx"))

    with io.open(os.path.join(folder, "settlement_v2-123456789-manual-2026-10-15.csv"), "w",
                 encoding="utf-8") as fh:
        fh.write("TRANSACTION_DATE;SETTLEMENT_DATE;SOURCE_ID;TRANSACTION_TYPE;"
                 "TRANSACTION_AMOUNT;FEE_AMOUNT;REAL_AMOUNT;TAXES_AMOUNT\n")
        for d, sid, t, a, f, real, tax in SETTLEMENT:
            fh.write("%s;%s;%s;%s;%.2f;%.2f;%.2f;%.2f\n" % (d, d, sid, t, a, f, real, tax))
    return folder


def _derive(folder):
    rb.set_site("pe")
    p0, p1 = rb.month_bounds("2026-09")
    with contextlib.redirect_stdout(io.StringIO()):
        return rb.derive("PE_TEST", folder, p0, p1)


def test_pe_numbers_bridges_and_checks(tmp_path):
    S = _derive(make_pe_folder(tmp_path / "PE_TEST" / "20261015_133021"))
    m = S["m"]
    assert S["cycle_day"] == 1                                  # 自然月账单，不推切期日
    assert [(y, mo) for y, mo, _a, _b in S["bill_cycles"]] == [(2026, 9)]
    assert round(m["ing"], 2) == 418.00
    # Flex 补贴并进运费收入，不再被倒推成"负的卖家优惠券"
    assert round(m["bonif"], 2) == 4.35 and m["bonif_n"] == 1
    assert round(m["env_ing"], 2) == 8.25 and round(m["env_ing_buyer"], 2) == 3.90
    assert abs(m["coupon"]) < 0.005 and m["coupon_n"] == 0
    assert abs(m["resid"]) < 0.005                              # 桥 A
    assert round(m["payable"], 2) == round(m["pay_auto"], 2) == 80.00   # 桥 B（Mi página 冲平）
    assert round(m["com_gross"], 2) == round(m["com_imp"], 2) == 48.16
    assert (m["com_match_n"], m["com_match_of"]) == (3, 3)
    assert round(m["iva_in_gmv"], 2) == round(418 * 0.18 / 1.18, 2)

    c = {x["key"]: x for x in rb.validate([S])}
    for key in ("bridge_a", "bridge_b", "fee_catalog", "bill_block_balance", "no_withholding",
                "billing_present", "sales_window"):
        assert c[key]["ok"] and not c[key]["skipped"], (key, c[key]["detail"])
    assert "tax_rate" not in c
    assert c["no_withholding"]["name"] == "平台未代扣税（秘鲁）"
    assert "IGV" in c["no_withholding"]["detail"] and "IVA" not in c["no_withholding"]["detail"]
    assert c["cargos_full"]["skipped"] and "秘鲁没有 Full" in c["cargos_full"]["detail"]

    # MercadoPago 逐单：运费（SETTLEMENT_SHIPPING）和 Flex 补贴（CASHBACK）是单独一笔，
    # 差额正好等于该单运费收入 —— 都被认出来
    MO = S["mp_orders"]
    assert (MO["linked"], MO["agree"], len(MO["differ"])) == (3, 1, 2)
    assert MO["ship_explained"][0] == 2 and round(MO["ship_explained"][1], 2) == 8.25
    assert "CASHBACK" in c["mp_settlement"]["detail"]


def test_negative_descuento_stays_a_seller_coupon(tmp_path):
    """负数不是补贴：照旧由 Total 倒推成卖家优惠券，不并进运费收入。"""
    orders = [ORDERS[0], ("2000000000000004", "5 de septiembre de 2026 09:00 hs.", 100, -12.0,
                          None, None, -10.0, 78.0, "SKU-C", "Correo y puntos de despacho")]
    S = _derive(make_pe_folder(tmp_path / "f", orders))
    m = S["m"]
    assert round(m["bonif"], 2) == 4.35
    assert round(m["coupon"], 2) == 10.00 and m["coupon_n"] == 1
    assert abs(m["resid"]) < 0.005


def test_pe_workbook_speaks_pen_igv_and_has_no_full(tmp_path):
    folder = make_pe_folder(tmp_path / "PE_TEST" / "20261015_133021")
    out = str(tmp_path / "pe.xlsx")
    with contextlib.redirect_stdout(io.StringIO()):
        # 不给 --site：从下载目录的 run_meta.json 读出 pe
        rc = rb.main(["--month", "2026-09", "--store", "PE_TEST=%s" % folder,
                      "--out", out, "--no-recalc", "--build-date", "2026-10-15"])
    assert rc == rb.EXIT_OK
    assert rb.SITE == "pe"
    wb = load_workbook(out)
    assert wb.sheetnames == rb.SHEET_NAMES
    assert "秘鲁" in wb[rb.SHEET_NAMES[0]]["A1"].value
    assert "PEN" in wb[rb.SHEET_NAMES[0]]["A2"].value

    pl = wb[rb.SHEET_NAMES[1]]
    rows = {pl.cell(r, 1).value: r for r in range(5, pl.max_row + 1) if pl.cell(r, 1).value}
    assert pl["C%d" % rows["商品销售收入 (GMV)"]].number_format == "#,##0.00;(#,##0.00);-"
    assert pl["C%d" % rows["运费收入"]].value == 8.25
    assert "Descuentos y bonificaciones" in pl["B%d" % rows["运费收入"]].value
    assert pl["C%d" % rows["减：卖家承担优惠券/折扣"]].value == 0
    assert not any("代扣代缴" in str(k) for k in rows)
    assert any("IGV" in str(k) for k in rows)                   # 备查项：GMV 中所含 IGV
    # 经营贡献毛利 = 418 + 8.25 − 48.16 − 12.7 − 80（广告；Mi página 已冲平）
    tot = sum(pl["C%d" % r].value for r in range(rows["商品销售收入 (GMV)"],
                                                 rows["减：账单级未登记科目 ⚠"] + 1)
              if isinstance(pl["C%d" % r].value, (int, float)))
    assert round(tot, 2) == 285.39

    for i in (5, 8):                                             # ⑥ 退货质检、⑨ 库存
        ws = wb[rb.SHEET_NAMES[i]]
        assert "秘鲁没有 Full 仓" in ws["A2"].value
        assert ws.max_row <= 4                                  # 只有说明，没有空表头

    gaps = wb[rb.SHEET_NAMES[9]]
    texts = [str(c.value) for row in gaps.iter_rows() for c in row if c.value]
    assert any("boleta" in t for t in texts)                     # B001-… 发票
    assert not any("缺少 Full 费用明细报表" in t for t in texts)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and not v.startswith("="):
                    for bad in ("智利", "MXN", "CLP", "比索", "SII", "切期"):
                        assert bad not in v, (ws.title, cell.coordinate, bad, v[:80])
                    assert not rb._IVA_RE.search(v), (ws.title, cell.coordinate, v[:80])


def test_pe_and_mx_folders_do_not_mix(tmp_path):
    pe = make_pe_folder(tmp_path / "PE_TEST" / "a")
    mx = tmp_path / "MX_TEST" / "b"
    mx.mkdir(parents=True)
    (mx / "run_meta.json").write_text('{"site": "mx"}')
    out = str(tmp_path / "x.xlsx")
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = rb.main(["--month", "2026-09", "--store", "PE_TEST=%s" % pe,
                      "--store", "MX_TEST=%s" % mx, "--out", out, "--no-recalc"])
    assert rc == rb.EXIT_CHECKS                     # 剔除的店铺记一条未通过的校验
    assert "[剔除] MX_TEST" in buf.getvalue()
