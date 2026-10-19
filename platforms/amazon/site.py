# -*- coding: utf-8 -*-
"""Per-country facts for Amazon. One entry per country the package serves
(core/markets.py PLATFORMS["AZ"].countries). Facts from discovery (FORMS.md)."""
from core import store_config

SITES = {
    "MX": {
        "name": "Amazon 墨西哥",
        "host": "https://sellercentral.amazon.com.mx",
        "locale": "es_MX",                 # #ngstrim-header data-locale
        "currency": "MXN",
        "decimals": 2,
        "vat": 0.16,                       # IVA
        # Amazon's monthly reports run in Central time ("GMT-6" on the page).
        "tz": "America/Mexico_City",
        "utc_offset": -6,
        # Seller fee invoices: Amazon's settlement periods are about two weeks.
        "invoice_period_days": 14,
    },
}


def get(country):
    try:
        return SITES[str(country).upper()]
    except KeyError:
        raise store_config.StoreConfigError(
            "亚马逊 还没有 %r 的站点资料（目前：%s）"
            % (country, ", ".join(sorted(SITES))))
