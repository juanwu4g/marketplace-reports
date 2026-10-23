# -*- coding: utf-8 -*-
"""The MercadoLibre workbook's module layout (platforms/mercadolibre/report/).

report_build.py is the entry point and keeps every name the old single file had;
the layers below it only import downwards; per-country values live in one place
(report/site.py) and are never copied; ②'s formula table names real rows.
The workbook itself is covered cell by cell by test_golden.py.
"""
import ast
import importlib
import os

import pytest

from platforms.mercadolibre import report_build as rb
from platforms.mercadolibre.report import site

PKG = os.path.join(os.path.dirname(os.path.abspath(rb.__file__)), "report")

# bottom to top: a module may import only from modules before it
LAYERS = ["site", "catalog", "loading", "derived", "checks", "sheet_common", "sheet_readme",
          "sheet_pl", "sheet_bridge", "sheet_fees", "sheet_returns", "sheet_sku", "sheet_trend",
          "sheet_inventory", "sheet_gaps", "figures"]

# names other code reads from report_build (tests, pl_billing_basis.py, test_golden.py)
USED_NAMES = """
set_site main run build_workbook build_context validate derive _chk _no_withholding_check
EXIT_OK EXIT_CHECKS EXIT_NO_STORE EXIT_STRICT_ABORT EXIT_CRASH
SITE SITE_PROFILES MNY CUR CUR_CN COUNTRY_CN VAT_NAME TAX_AUTH IVA_RATE WITHHOLDING HAS_FULL
BONIF_COL TOL_CENT CN_TOL BRIDGE_A_TOL BILL_CYCLE_DAY VENTAS_SHEET VENTAS_GLOB
VENTAS_HEADER_ROW BILLING_HEADER_ROW PAGOS_HEADER_ROW FILE_MONTH_TOKENS
V_ORDER_ID V_DATE V_STATUS V_UNITS V_REVENUE V_COMMISSION V_SHIP_INCOME V_SHIP_COST V_REFUND
V_TOTAL V_SKU V_TITLE V_PACK_FLAG
F_INVOICE F_CHARGE_DATE F_CHARGE_NO F_DETALLE F_CHANNEL F_STATE F_REVERSES F_AMOUNT F_SALE_NO
F_PACK_NO F_SALE_DATE
FEE_CATALOG FEE_BY_ES MONEY_REFUNDED MONEY_KEPT SHEET_NAMES
F fill fmt money BOX PCT C_GREY C_NAVY C_SUB C_RED C_WARN C_INPUT C_CALC
localize _IVA_RE bill_cycles infer_cycle_day month_bounds file_month_token parse_spanish_date
load_ventas load_billing load_returns _ventas_header_row mp_by_order
build_pl_rows pl_value pl_memo_no_withholding sheet_pl
""".split()


def module_source(name):
    with open(os.path.join(PKG, name + ".py"), encoding="utf-8") as fh:
        return fh.read()


def test_every_module_is_in_the_layer_list():
    on_disk = sorted(f[:-3] for f in os.listdir(PKG) if f.endswith(".py") and f != "__init__.py")
    assert on_disk == sorted(LAYERS)


@pytest.mark.parametrize("name", LAYERS)
def test_layers_only_import_downwards(name):
    for node in ast.walk(ast.parse(module_source(name))):
        if isinstance(node, ast.ImportFrom) and (node.module or "").startswith("platforms."):
            assert node.module.startswith("platforms.mercadolibre.report"), node.module
            if node.module == "platforms.mercadolibre.report":
                deps = [a.name for a in node.names]
            else:
                deps = [node.module.rsplit(".", 1)[1]]
            for d in deps:
                assert LAYERS.index(d) < LAYERS.index(name), "%s imports %s" % (name, d)
        if isinstance(node, ast.Import):
            assert not any(a.name.startswith("platforms.") for a in node.names)


def test_report_build_keeps_the_names_others_use():
    missing = [n for n in USED_NAMES if not hasattr(rb, n)]
    assert not missing


@pytest.mark.parametrize("code", ["cl", "pe", "mx"])
def test_country_values_follow_set_site_everywhere(code):
    try:
        assert rb.set_site(code) == code
        prof = dict(site.SITE_DEFAULTS, **site.SITE_PROFILES[code])
        for key, value in prof.items():
            assert getattr(rb, key) == value == getattr(site, key)
        assert rb.SITE == code
        assert rb.V_REVENUE == "Ingresos por productos (%s)" % prof["CUR"]
    finally:
        rb.set_site("mx")


@pytest.mark.parametrize("name", LAYERS[1:])
def test_no_module_keeps_its_own_copy_of_a_country_value(name):
    """A module that did `from ...site import MNY` would keep Mexico's format in a
    Chilean workbook. Country values are read as site.NAME only."""
    mod = importlib.import_module("platforms.mercadolibre.report." + name)
    assert not set(vars(mod)) & set(site.STATE)
    assert not set(vars(rb)) & set(site.STATE)          # report_build serves them live


@pytest.mark.parametrize("code", ["mx", "cl", "pe"])
def test_pl_formula_table_names_real_rows(code):
    from core import report_family as RF
    from platforms.mercadolibre.report import sheet_pl
    try:
        rb.set_site(code)
        rows = sheet_pl.build_pl_rows()
        labels = [r[1] for r in rows]
        assert len(labels) == len(set(labels))                 # formulas refer to rows by label
        kinds = dict((r[1], r[3]) for r in rows)
        formula_rows = [label for label, _f in sheet_pl.PL_FORMULAS]
        assert sorted(formula_rows) == sorted(l for l, k in kinds.items() if k in ("sum", "f"))
        for label, formula in sheet_pl.PL_FORMULAS:
            refs = [formula.first, formula.last] if isinstance(formula, RF.Span) else \
                [k.lstrip("-") for k in formula]
            for ref in refs:
                assert ref in kinds, (label, ref)
                assert labels.index(ref) < labels.index(label), (label, ref)
    finally:
        rb.set_site("mx")
