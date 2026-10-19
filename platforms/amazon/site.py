# -*- coding: utf-8 -*-
"""Per-country facts for Amazon. One entry per country the package serves
(core/markets.py PLATFORMS["AZ"].countries). Fill in from discovery."""
from core import store_config

SITES = {
    "MX": {
        "name": "Amazon 墨西哥",
        "host": "https://TODO",            # seller centre, no trailing slash
        "locale": "TODO",                  # the UI language the routes are written for
        "currency": "TODO",
        "decimals": 2,                     # money decimals of the currency (CLP: 0)
        "tz": "TODO/Timezone",             # IANA name, for site_time
        "utc_offset": 0,                   # fallback when tzdata is missing
    },
}


def get(country):
    try:
        return SITES[str(country).upper()]
    except KeyError:
        raise store_config.StoreConfigError(
            "亚马逊 还没有 %r 的站点资料（目前：%s）"
            % (country, ", ".join(sorted(SITES))))
