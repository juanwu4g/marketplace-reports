# -*- coding: utf-8 -*-
"""Reventa (MercadoLibre buys the item from the seller and resells it) in both
MercadoLibre report bases. Synthetic folders only: the bill-basis Mexican folder
of test_ml_bill_report, with a sales export in the format MercadoLibre has used
since Reventa arrived (2026-10): a "Modelo de venta" column, the unit-price
header renamed, and Reventa rows that carry the net amount MercadoLibre pays,
no charge, no shipping, no bill row.
"""
import contextlib
import glob
import io
import os

import pytest
from openpyxl import load_workbook

from core import report_kit as K
from platforms.mercadolibre import report_build as rb
from platforms.mercadolibre.bill_report import build, model, readers
from platforms.mercadolibre.bill_report import reventa as RV_BILL
from platforms.mercadolibre.report import reventa as RV_ORDER
from tests import test_ml_bill_report as BT

OLD_PRICE = "Precio unitario de venta de la publicación (MXN)"
NEW_PRICE = "Precio unitario para venta directa / Precio sugerido para reventa (MXN)"

# sale no, sale date, state, units, net paid by ML, SKU, listing, title, suggested price, refund
REVENTA_ROWS = [
    ("2001", "20 de septiembre de 2026 10:00 hs.", "Entregado", 2, 150.0, "SKU-A", "MLM111", "Bocina A", 100.0,
     None),
    ("2002", "21 de septiembre de 2026 11:00 hs.", "Cancelada por el comprador", 1, 220.0, "SKU-B", "MLM222",
     "Bocina B", 300.0, -220.0),
    ("2003", "2 de octubre de 2026 09:00 hs.", "Entregado", 1, 75.0, "SKU-A", "MLM111", "Bocina A", 100.0, None),
]


def add_reventa(folder):
    """Rewrite the folder's sales export in the new format, with REVENTA_ROWS."""
    path = glob.glob(os.path.join(folder, "*Ventas*.xlsx"))[0]
    wb = load_workbook(path)
    ws = wb.active
    hdr = dict((c.value, c.column) for c in ws[6] if c.value)
    ws.cell(row=6, column=hdr[OLD_PRICE], value=NEW_PRICE)
    hdr[NEW_PRICE] = hdr.pop(OLD_PRICE)
    for name in ("Modelo de venta", "Anulaciones y reembolsos (MXN)", "Total (MXN)"):
        hdr[name] = ws.max_column + 1
        ws.cell(row=6, column=hdr[name], value=name)
    last = ws.max_row
    for r in range(7, last + 1):
        ws.cell(row=r, column=hdr["Modelo de venta"], value="Venta directa")
    for i, (no, day, state, units, net, sku, listing, title, price, refund) in enumerate(REVENTA_ROWS, last + 1):
        for name, v in (("# de venta", no), ("Fecha de venta", day), ("Estado", state),
                        ("Paquete de varios productos", "No"), ("Unidades", units),
                        ("Ingresos por productos (MXN)", net), ("SKU", sku), ("# de publicación", listing),
                        ("Título de la publicación", title), (NEW_PRICE, price), ("Modelo de venta", "Reventa"),
                        ("Anulaciones y reembolsos (MXN)", refund), ("Total (MXN)", net + (refund or 0.0))):
            ws.cell(row=i, column=hdr[name], value=v)
    wb.save(path)
    return folder


@pytest.fixture
def plain(tmp_path):
    return BT.make_mx_folder(tmp_path / "plain")


@pytest.fixture
def reventa(tmp_path):
    return add_reventa(BT.make_mx_folder(tmp_path / "rv"))


def test_the_two_copies_of_the_module_are_identical():
    """Each basis has its own copy (they never import each other): keep them the same."""
    a, b = (io.open(m.__file__.replace(".pyc", ".py"), encoding="utf-8").read() for m in (RV_BILL, RV_ORDER))
    assert a == b


def test_header_facts():
    assert RV_BILL.is_reventa("Reventa") and RV_BILL.is_reventa(" reventa ")
    assert not RV_BILL.is_reventa("Venta directa") and not RV_BILL.is_reventa(None)
    assert RV_BILL.unit_price_column(["x", OLD_PRICE]) == OLD_PRICE
    assert RV_BILL.unit_price_column(["x", NEW_PRICE]) == NEW_PRICE
    assert RV_BILL.unit_price_column(["x"]) is None


def test_renamed_unit_price_header_is_still_read(reventa):
    V, facts = readers.read_ventas(glob.glob(os.path.join(reventa, "*Ventas*.xlsx"))[0])
    by = dict((r.sale_no, r) for r in V.itertuples(index=False))
    assert by["1001"].unit_price == 100.0 and not by["1001"].reventa
    assert by["2001"].reventa and by["2001"].unit_price == 100.0
    assert facts["reventa"] == 3


# ------------------------------------------------------------------ bill basis

def test_bill_basis_takes_reventa_sales_from_the_export(plain, reventa):
    base, S = BT.derive(plain), BT.derive(reventa)
    b, pl = base["m"]["pl"], S["m"]["pl"]
    rv = S["m"]["reventa"]
    # in September: 2001 (150) and 2002 (220, cancelled); 2003 is October's
    assert (rv["n"], rv["units"], rv["sales"], rv["refund"], rv["n_refund"]) == (2, 3, 370.0, -220.0, 1)
    assert rv["suggested"] == 2 * 100.0 + 300.0
    assert rv["direct_gmv"] == pytest.approx(b["gmv_billed"])
    assert str(rv["first"]) == "2026-09-20"
    assert pl["gmv_billed"] == pytest.approx(b["gmv_billed"] + 370.0)
    assert pl["rev_cancel"] == pytest.approx(b["rev_cancel"] - 220.0)
    assert pl["net_gmv"] == pytest.approx(b["net_gmv"] + 150.0)
    assert pl["contrib"] == pytest.approx(b["contrib"] + 150.0)
    # nothing else moves: no fee, no buyer shipping, no withholding on a Reventa sale
    for k in ("fees_total", "ship_income", "withheld", "seller_discount", "platform_subsidy"):
        assert pl[k] == pytest.approx(b[k]), k
    assert S["outside"]["2001"]["wh_src"] == model.WH_NA
    # the sale goes to its SKU with its units; the cancelled one keeps no cost
    assert S["sales"]["2001"]["sku"] == "SKU-A" and S["sales"]["2001"]["source"] == model.REVENTA
    assert S["sku"]["SKU-A"]["units"] == base["sku"]["SKU-A"]["units"] + 2
    assert S["sku"]["SKU-B"]["units_cancel"] == base["sku"]["SKU-B"]["units_cancel"] + 1
    assert all(abs(v) < 1e-6 for v in model.residuals(S).values())
    checks = dict((c["key"], c) for c in model.validate(S))
    assert checks["reventa"]["ok"] and "2 笔" in checks["reventa"]["detail"]


def test_bill_basis_without_reventa_is_untouched(plain):
    S = BT.derive(plain)
    assert S["m"]["reventa"]["n"] == 0
    assert "reventa" not in [c["key"] for c in model.validate(S)]


def test_bill_basis_leaves_a_reventa_sale_the_bill_charges_to_the_bill(reventa):
    """A sale the bill has a commission row for is billed, whatever the export says."""
    path = glob.glob(os.path.join(reventa, "*Ventas*.xlsx"))[0]
    wb = load_workbook(path)
    ws = wb.active
    col = next(c.column for c in ws[6] if c.value == "Modelo de venta")
    row = next(r for r in range(7, ws.max_row + 1) if str(ws.cell(row=r, column=1).value) == "1001")
    ws.cell(row=row, column=col, value="Reventa")
    wb.save(path)
    S = BT.derive(reventa)
    assert S["sales"]["1001"]["source"] == "commission"
    assert S["reventa"]["billed"] == ["1001"]
    assert not dict((c["key"], c) for c in model.validate(S))["reventa"]["ok"]


def test_bill_workbook_shows_reventa_only_when_there_is_some(plain, reventa, tmp_path):
    labels = {}
    for name, folder in (("plain", plain), ("rv", reventa)):
        res = build.run("STORE_MX", folder, "2026-09", "mx", out_dir=str(tmp_path / name), recalc=False)
        wb = load_workbook(res["out"])
        assert wb.sheetnames == K.SHEET_NAMES
        labels[name] = ([c.value for c in wb[K.SHEET_NAMES[1]]["A"] if c.value],
                        [c.value for c in wb[K.SHEET_NAMES[6]]["A"] if c.value])
    assert not any("Reventa" in str(x) for x in labels["plain"][0] + labels["plain"][1])
    assert "其中：Reventa 销售（平台支付净额，已含在商品销售额中）" in labels["rv"][0]
    assert any(str(x).startswith("表三") for x in labels["rv"][1])
    # the rows above the new memo lines are where they were
    n = len(labels["plain"][0])
    assert labels["rv"][0][:n] == labels["plain"][0]


# ------------------------------------------------------------------ order basis

def _order(folder):
    # the bill-basis fixture's storage detail is not in the layout the order basis reads
    for f in glob.glob(os.path.join(folder, "*Costos_por_servicio_almacenamiento.xlsx")):
        os.remove(f)
    rb.set_site("mx")
    p0, p1 = rb.month_bounds("2026-09")
    with contextlib.redirect_stdout(io.StringIO()):
        return rb.derive("STORE_MX", folder, p0, p1)


def test_order_basis_splits_reventa_from_direct_sales(plain, reventa):
    base, S = _order(plain), _order(reventa)
    m, rv = S["m"], S["m"]["rv"]
    assert (rv["n"], rv["units"], rv["ing"], rv["refund"], rv["suggested"]) == (2, 3.0, 370.0, -220.0, 500.0)
    assert m["ing"] == pytest.approx(base["m"]["ing"] + 370.0)
    assert m["ing_direct"] == pytest.approx(base["m"]["ing"])
    # the withholding rate is measured on direct sales only: Reventa does not dilute it
    assert m["tax_ret"] == pytest.approx(base["m"]["tax_ret"])
    assert m["tax_rate_gross"] == pytest.approx(base["m"]["tax_rate_gross"])
    assert base["m"]["rv"]["n"] == 0 and base["m"]["ing_direct"] == base["m"]["ing"]
    with contextlib.redirect_stdout(io.StringIO()):
        keys = [c["key"] for c in rb.validate([S])]
        base_keys = [c["key"] for c in rb.validate([base])]
    assert "reventa" in keys and "reventa" not in base_keys
    from platforms.mercadolibre.report import sheet_pl, sheet_sku
    assert [k for k, _ in sheet_pl.pl_memo([S])][-4:] == [
        "reventa_sales", "reventa_refund", "direct_gmv", "reventa_suggested"]
    assert [k for k, _ in sheet_pl.pl_memo([base])] == list(sheet_pl.MEMO_KEYS)
    table, skipped = sheet_sku.reventa_compare(S)
    assert table["SKU-A"]["r_units"] == 2 and table["SKU-A"]["r_net"] == 150.0
    assert table["SKU-A"]["r_list"] == 200.0 and skipped >= 1          # 2002 was refunded
