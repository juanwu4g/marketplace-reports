# -*- coding: utf-8 -*-
"""Per-country facts for Liverpool. One entry per country the package serves
(core/markets.py PLATFORMS["LV"].countries). Facts from discovery (FORMS.md)."""
from core import store_config

SITES = {
    "MX": {
        "name": "Liverpool 墨西哥",
        # The marketplace seller centre (orders, logistic costs, account balance),
        # opened with the store's marketplace user (config: marketplace_store).
        "host": "https://marketplace.liverpool.com.mx",
        # The supplier portal (SAP NetWeaver): the ledger, the sales report, the
        # payments and the fee CFDIs, opened with the supplier user (the store itself).
        "portal": "https://proveedores.liverpool.com.mx",
        # Opening the single-sign-on portal first gives the browser the Akamai
        # cookies without which the supplier portal answers "Access Denied".
        "warmup": "https://entradaunica.liverpool.com.mx/portal",
        "locale": "es_MX",
        "currency": "MXN",
        "decimals": 2,
        "vat": 0.16,                       # IVA
        "tz": "America/Mexico_City",
        "utc_offset": -6,
    },
}


def get(country):
    try:
        return SITES[str(country).upper()]
    except KeyError:
        raise store_config.StoreConfigError(
            "利物浦 还没有 %r 的站点资料（目前：%s）"
            % (country, ", ".join(sorted(SITES))))
