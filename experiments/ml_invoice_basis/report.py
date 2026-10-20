# -*- coding: utf-8 -*-
"""One store, one month, one download folder -> every number of the experiment.

compute() switches the MercadoLibre report to the folder's site (run_meta.json),
runs the current basis (derive) and the invoice basis on the same invoice rows,
and returns plain data; the site is restored before it returns, so the workbook
writers never depend on module globals."""

import contextlib
import io

from core import report_kit as K
from platforms.mercadolibre.report import site

from experiments.ml_invoice_basis import accounting, compare, invoice, vat
from experiments.ml_invoice_basis import mercadopago as MP
from experiments.ml_invoice_basis.common import ExperimentError, folder_site


def compute(store, folder, month, quiet=True, cost_table=None):
    """cost_table: (table, note) to use instead of the configured SKU cost table
    (tests); None reads core.sku_costs.load_configured()."""
    with folder_site(folder) as code:
        buf = io.StringIO()
        with (contextlib.redirect_stdout(buf) if quiet else contextlib.nullcontext()):
            try:
                S = compare.current_basis(store, folder, month)
            except IOError as e:
                raise ExperimentError("现行口径（derive）读不了这个目录：%s" % e)
            ctx = invoice.load(folder, month, B=S["billing_all"])
            period, cal = invoice.views(ctx)
            invs = {"period": invoice.analyze(period, ctx["B"]),
                    "calendar": invoice.analyze(cal, ctx["B"])}
            orders, parent_of = compare.order_index(S["ventas_all"])
            where = compare.billed_where(ctx["B"], orders, parent_of)
            recs = dict((v, compare.reconcile(invs[v], orders, parent_of, month, where, ctx["B"]))
                        for v in invs)
            lines = compare.comparison(S, invs, recs, ctx["cycle_day"])
            unavailable = compare.unavailable(S)
            D = {"store": store, "folder": folder, "month": month, "site": code, "S": S,
                 "ctx": ctx, "docs": ctx["docs"], "invs": invs, "recs": recs, "lines": lines,
                 "orders": orders, "parent_of": parent_of, "dec": site.MONEY_DEC}
            D["mp"] = MP.load(folder)
            D["links"] = accounting.links(ctx["B"])
            D["vat"] = vat.policy(code)
            D["vat_ev"] = vat.evidence(folder, ctx["B"], D["mp"]["rows"], site.CN_TOL)
            D["vat_verdict"], D["vat_lines"] = vat.conclusion(D["vat_ev"], D["vat"])
            table, table_note = (cost_table if cost_table is not None
                                 else accounting.load_cost_table())
            D["cost_note"] = table_note
            D["acc"] = dict((v, accounting.reading(invs[v], D, D["mp"], D["links"], table))
                            for v in invs)
            D["dates"] = accounting.three_dates(D)
            D["cash"] = accounting.cash_bridge(D, "calendar", D["acc"]["calendar"])
            D["coupon_check"] = accounting.coupon_check(D, D["mp"], D["links"])
            D["withholding"] = site.WITHHOLDING
        return dict(D, **{
            "store": store, "folder": folder, "month": month, "site": code,
            "country": site.COUNTRY_CN, "cur": site.CUR, "dec": site.MONEY_DEC,
            "money": K.money_format(site.MONEY_DEC), "withholding": site.WITHHOLDING,
            "cycle_day": ctx["cycle_day"], "docs": ctx["docs"], "S": S, "invs": invs,
            "recs": recs, "lines": lines, "unavailable": unavailable, "orders": orders,
            "parent_of": parent_of,
            "log": buf.getvalue() if quiet else "",
        })


def key_numbers(D):
    """The headline numbers of one snapshot, for the stability comparison and the
    console. Order basis from derive(); invoice basis from the two readings."""
    m, S = D["S"]["m"], D["S"]
    out = {"order": {
        "count": m["orders_unique"], "units": m["units"], "gmv": m["ing"],
        "refund": -m["refund"], "ship_in": m["env_ing"], "coupon": -m["coupon"],
        "commission": -m["com_gross"], "shipping": -m["env_cost"], "returns": -m["dev_net"],
        "tax": m["tax_ret"] if D["withholding"] else None,
        "contrib": S["contrib"]}}
    for es, v in m["bill_fee_net"].items():
        out["order"]["fee:" + es] = -v
    out["order"]["unc"] = -(m["unc_order"] + m["unc_bill"])
    out["order"]["fees"] = (out["order"]["commission"] + out["order"]["shipping"]
                            + out["order"]["returns"] - m["unc_order"]
                            - sum(m["bill_fee_net"].values()) - m["unc_bill"])
    for v, inv in D["invs"].items():
        n, bb = inv["n"], inv["by_bucket"]
        d = {"count": n["n_sales"], "units": n["units"], "gmv": n["gross_sales"],
             "refund": -n["rev_sales"], "ship_in": n["ship_income"] - n["ship_income_rev"],
             "coupon": None,
             "commission": -bb.get("commission", {}).get("net", 0.0),
             "shipping": -bb.get("shipping", {}).get("net", 0.0),
             "returns": -bb.get("returns", {}).get("net", 0.0), "tax": None,
             "contrib": n["contribution"], "fees": -n["fee_total"],
             "unc": -(n["unk_order"] + n["unk_bill"])}
        for f in inv["fees"]:
            if not f["order_level"]:
                d["fee:" + f["es"]] = -f["net"]
        a = D["acc"][v]
        have = a["mp"]["available"]
        d.update({"refund_actual": -a["rev"]["value"],
                  "refund_err": a["rev"]["err"] if have and a["rev"]["n_known"] else None,
                  "coupon_mp": -a["coupon"] if a["mp"]["coupon_available"] else None,
                  "accrual": -a["accrual"]["total"], "contrib_adj": a["contrib_adj"],
                  "net_sales_exvat": a["vat"]["rev_net"], "contrib_exvat": a["vat"]["contrib_net"]})
        out[v] = d
    # the order basis on the same net-of-VAT rule (revenue lines ÷ (1 + rate); fees too
    # when their VAT is creditable)
    r, credit = D["vat"]["rate"], D["vat"]["credit"]
    rev_keys = ("gmv", "refund", "cn_adj", "ship_in", "coupon")
    money = [ln for ln in D["lines"] if ln["kind"] == "m"]
    rev = sum(ln["cur"] or 0.0 for ln in money if ln["key"] in rev_keys)
    fees = sum(ln["cur"] or 0.0 for ln in money if ln["key"] not in rev_keys)
    out["order"].update({"refund_actual": None, "refund_err": None, "coupon_mp": None,
                         "accrual": None, "contrib_adj": None,
                         "net_sales_exvat": rev / (1 + r),
                         "contrib_exvat": rev / (1 + r) + (fees / (1 + r) if credit else fees)})
    out["mp_source"] = D["mp"]["source"]
    return out
