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

from experiments.ml_invoice_basis import compare, invoice
from experiments.ml_invoice_basis.common import ExperimentError, folder_site


def compute(store, folder, month, quiet=True):
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
        return {
            "store": store, "folder": folder, "month": month, "site": code,
            "country": site.COUNTRY_CN, "cur": site.CUR, "dec": site.MONEY_DEC,
            "money": K.money_format(site.MONEY_DEC), "withholding": site.WITHHOLDING,
            "cycle_day": ctx["cycle_day"], "docs": ctx["docs"], "S": S, "invs": invs,
            "recs": recs, "lines": lines, "unavailable": unavailable, "orders": orders,
            "parent_of": parent_of,
            "log": buf.getvalue() if quiet else "",
        }


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
        out[v] = d
    return out
