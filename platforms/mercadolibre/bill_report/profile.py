# -*- coding: utf-8 -*-
"""Country facts this report needs (Mexico / Chile / Peru) — its own table.

The IVA rule is the one the official MercadoLibre report documents for each
country (root README §5, platforms/mercadolibre/README.md "tax" row): amounts
are shown **with IVA, as MercadoLibre states them**; the IVA they contain is a
memo line, never deducted in the P&L. Mexico's platform withholds IVA/ISR
through MercadoPago — that is not in the invoices, so the workbook says "not in
this baseline". Chile and Peru: nothing withheld.
"""
import collections

Profile = collections.namedtuple(
    "Profile", "site country country_cn currency currency_cn decimals vat_rate vat_name tax_auth "
               "withholding has_full listing_prefix ventas_sheet tol")

PROFILES = {
    "mx": Profile("mx", "MX", "墨西哥", "MXN", "墨西哥比索", 2, 0.16, "IVA", "SAT",
                  True, True, "MLM", "Ventas MX", 0.01),
    "cl": Profile("cl", "CL", "智利", "CLP", "智利比索", 0, 0.19, "IVA", "SII",
                  False, True, "MLC", "Ventas CL", 1.0),
    "pe": Profile("pe", "PE", "秘鲁", "PEN", "秘鲁索尔", 2, 0.18, "IGV", "SUNAT",
                  False, False, "MLP", "Ventas PE", 0.01),
}


def get(site):
    code = str(site or "").strip().lower()
    if code not in PROFILES:
        raise ValueError("未知站点 %r（可用：%s）" % (site, ", ".join(sorted(PROFILES))))
    return PROFILES[code]


def vat_in(amount, prof):
    """The IVA / IGV contained in an amount that includes it: amount × r / (1 + r)."""
    return amount * prof.vat_rate / (1.0 + prof.vat_rate)
