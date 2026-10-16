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
builder - see docs/PLATFORM_PLAYBOOK.zh.md section 9.

Deliberately dependency-free, like store_config.py.
"""
import re

from core import store_config

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
    "pe": {
        "name": "秘鲁",
        "site_id": "MPE",
        "seller_host": "https://vendedores.mercadolibre.com.pe",
        "mp_host": "https://www.mercadopago.com.pe",
        "locale": "es_PE",
        "currency": "PEN",
        "platform_hint": "秘鲁",
        "tz": "America/Lima", "utc_offset": -5,
        # 秘鲁把九月写成 "Setiembre"（账单卡片、明细页标题、日期 "30/set/2026"）
        "september": "setiembre",
        # 销售 Excel 的下载通知过一阵就清空：申请完必须当场取（run_downloads）
        "ventas_collect_now": True,
        # 秘鲁没有 Full 仓：库存报表页（space_management）是 404，不去下载
        "has_full": False,
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
    """Resolve a store to its site code ("mx" / "cl" / "pe")."""
    try:
        configured = store_config.configured_country(store_name)
    except store_config.StoreConfigError:
        raise
    except Exception:
        configured = None       # 没有 config.json（测试、新机器）：按下面的规则判断
    if configured:
        code = str(configured).lower()
        get(code)
        return code
    for code, facts in SITES.items():
        if platform_name and facts["platform_hint"] in platform_name:
            return code
    m = re.search(r"(?:^|[_\-\s])(MX|CL|PE)(?:$|[_\-\s(（])", store_name or "", re.I)
    if m:
        return m.group(1).lower()
    return DEFAULT
