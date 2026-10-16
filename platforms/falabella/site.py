# -*- coding: utf-8 -*-
"""Per-country facts for Falabella. Only Chile today; Peru and Colombia run on
the same seller centre (the payments page has a CHILE / PERU / COLOMBIA switch)
and get an entry here when a store needs them."""
from core import store_config

SITES = {
    "CL": {
        "name": "Falabella Chile",
        "host": "https://sellercenter.falabella.com",
        "locale": "es-CL",             # data-localefull on the micro-frontends
        "currency": "CLP",
        "decimals": 0,
        "tz": "America/Santiago",
        "utc_offset": -3,
        "vat": 0.19,
        # Factura falabella.com periods run from this day to the day before it
        # in the next month ("21 agosto 2026 - 20 septiembre 2026").
        "billing_cycle_day": 21,
        # Lista de Órdenes: after choosing a start day only 13 more are selectable.
        "orders_max_days": 14,
    },
}


def get(country):
    try:
        return SITES[str(country).upper()]
    except KeyError:
        raise store_config.StoreConfigError(
            "Falabella 还没有 %r 的站点资料（目前：%s）"
            % (country, ", ".join(sorted(SITES))))
