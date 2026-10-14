# -*- coding: utf-8 -*-
"""
Which MercadoLibre country site a store belongs to.

The download layer was written against Mexico and had the host names and the
locale baked in as module constants. Chile runs the same seller centre on a
different host with a different locale and currency, so those few facts live
here and `meli_forms.set_site()` / `mercadopago.set_site()` rebind their
constants from it before a store's routes run.

A store's site is resolved, in order:

  1. config.json  ->  "stores": {"regions": {"Tecnohub_CL": "CL"}}
  2. the Ziniao platform_name of the store ("MercadoLibre-智利-本土")
  3. a _CL / _MX marker in the store name
  4. DEFAULT ("mx") - what every store was before sites existed

Only facts that are true of the *site* belong here. Accounting rules (fee
catalogue, tax treatment) are per-country too, but they live with the report
builder - see PLATFORM_PLAYBOOK.md section 9.

Deliberately dependency-free, like store_config.py.
"""
import re

import store_config

DEFAULT = "mx"

SITES = {
    "mx": {
        "name": "墨西哥",
        "site_id": "MLM",
        "seller_host": "https://vendedores.mercadolibre.com.mx",
        "mp_host": "https://www.mercadopago.com.mx",
        "locale": "es_MX",
        "currency": "MXN",
        "platform_hint": "墨西哥",
        "tz": "America/Mexico_City", "utc_offset": -6,
    },
    "cl": {
        "name": "智利",
        "site_id": "MLC",
        "seller_host": "https://vendedores.mercadolibre.cl",
        "mp_host": "https://www.mercadopago.cl",
        "locale": "es_CL",
        "currency": "CLP",
        "platform_hint": "智利",
        "tz": "America/Santiago", "utc_offset": -3,
    },
}


def get(code):
    """The facts for one site. Unknown codes raise - never fall back silently:
    running a Chilean store against the Mexican host fails in confusing ways."""
    try:
        return SITES[code]
    except KeyError:
        raise store_config.StoreConfigError(
            "未知站点 %r，可用：%s" % (code, ", ".join(sorted(SITES))))


def of_store(store_name, platform_name=None):
    """Resolve a store to its site code ("mx" / "cl")."""
    try:
        regions = store_config._stores().get("regions")
    except Exception:
        regions = None          # 没有 config.json（测试、新机器）：按下面的规则判断
    if isinstance(regions, dict) and store_name in regions:
        code = str(regions[store_name]).lower()
        get(code)
        return code
    for code, facts in SITES.items():
        if platform_name and facts["platform_hint"] in platform_name:
            return code
    m = re.search(r"(?:^|[_\-\s])(MX|CL)(?:$|[_\-\s(（])", store_name or "", re.I)
    if m:
        return m.group(1).lower()
    return DEFAULT
