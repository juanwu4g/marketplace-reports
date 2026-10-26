# -*- coding: utf-8 -*-
"""
Which platforms and countries exist — the registry every other module asks.

A **market** is one platform in one country: "ML-MX" (MercadoLibre Mexico),
"FA-CL" (Falabella Chile). Each store belongs to exactly one market. The runner,
the window, the alert files and the report names all work per market, so adding
a platform or a country is one entry here plus the platform's package — no
other module grows a list of names.

    PLATFORMS   code -> Platform (package, Chinese name, how Ziniao names it)
    COUNTRIES   code -> Chinese name (ISO 3166-1 alpha-2)

Pure data, no I/O: imported by tests, the window and the reports alike.
"""
import collections
import re

Platform = collections.namedtuple(
    "Platform", "code package name_zh name_en ziniao_pattern countries")

# `ziniao_pattern` matches the platform name Ziniao reports for a store
# ("MercadoLibre-墨西哥-本土", "falabella-智利"). It is a second, independent
# check that a store really is on this platform — the first is config.json.
# `countries` are the countries the package can handle; a store configured for
# another country is refused instead of being run against the wrong site.
PLATFORMS = collections.OrderedDict([
    ("ML", Platform("ML", "platforms.mercadolibre", "美客多", "MercadoLibre",
                    re.compile(r"mercado\s*libre", re.I), ("MX", "CL", "PE"))),
    ("FA", Platform("FA", "platforms.falabella", "Falabella", "Falabella",
                    re.compile(r"falabella", re.I), ("CL",))),
    # Ziniao names Amazon's platforms in Chinese: "亚马逊-墨西哥".
    ("AZ", Platform("AZ", "platforms.amazon", "亚马逊", "Amazon",
                    re.compile(r"amazon|亚马逊", re.I), ("MX",))),
    # Liverpool's stores are Ziniao "custom platform" stores (GENERIC_ZINIAO).
    ("LV", Platform("LV", "platforms.liverpool", "利物浦", "Liverpool",
                    re.compile(r"liverpool|^自定义平台$", re.I), ("MX",))),
])

# Ziniao's name for a store whose platform it does not know ("custom platform").
# It says nothing about the platform: a pattern may accept it (Liverpool's does),
# but it never suggests one (platform_from_ziniao).
GENERIC_ZINIAO = re.compile(r"^自定义平台$")

COUNTRIES = collections.OrderedDict([
    ("MX", "墨西哥"),
    ("CL", "智利"),
    ("PE", "秘鲁"),
    ("CO", "哥伦比亚"),
])

# Codes appear in file names finance keeps for years (docs/REPORT_FAMILY.md §1):
# once used, a code never changes.
#
# Stores that config.json does not describe belong here. Everything before the
# second platform was MercadoLibre Mexico, and production configs still rely on
# it — see store_config for when this default is NOT applied.
DEFAULT_PLATFORM = "ML"
DEFAULT_COUNTRY = "MX"

SEP = "-"


class UnknownMarket(ValueError):
    """A platform or country code nobody registered."""


def platform(code):
    try:
        return PLATFORMS[str(code or "").upper()]
    except KeyError:
        raise UnknownMarket("未知平台 %r，可用：%s" % (code, "、".join(PLATFORMS)))


def country_name(code):
    """MX -> 墨西哥. Unknown codes come back unchanged (labels must never crash)."""
    return COUNTRIES.get(str(code or "").upper(), str(code or ""))


def market_id(platform_code, country):
    return "%s%s%s" % (str(platform_code).upper(), SEP, str(country).upper())


def split(market):
    """'FA-CL' -> ('FA', 'CL'). A bare country 'CL' means MercadoLibre (legacy)."""
    text = str(market or "").strip().upper()
    if SEP in text:
        p, c = text.split(SEP, 1)
    else:
        p, c = DEFAULT_PLATFORM, text or DEFAULT_COUNTRY
    return p, c


def check(platform_code, country):
    """Raise UnknownMarket unless this platform is registered and serves this country."""
    p = platform(platform_code)
    c = str(country or "").upper()
    if c not in COUNTRIES:
        raise UnknownMarket("未知国家 %r，可用：%s" % (country, "、".join(COUNTRIES)))
    if c not in p.countries:
        raise UnknownMarket("%s 还不支持 %s（目前支持：%s）"
                            % (p.name_zh, country_name(c), "、".join(
                                country_name(x) for x in p.countries)))
    return p.code, c


def normalize(market):
    """Any spelling ('fa-cl', 'CL') -> canonical id ('FA-CL', 'ML-CL'), validated."""
    return market_id(*check(*split(market)))


def label(market):
    """'FA-CL' -> 'Falabella 智利'; 'ML-MX' -> '美客多 墨西哥'."""
    p, c = split(market)
    try:
        name = platform(p).name_zh
    except UnknownMarket:
        name = p
    return "%s %s" % (name, country_name(c))


def all_markets():
    """Every market the code supports, platform by platform, in registry order."""
    return [market_id(p.code, c) for p in PLATFORMS.values() for c in p.countries]


def ziniao_matches(platform_code, ziniao_platform_name):
    """Does Ziniao's platform name for a store agree with the configured platform?

    None / empty (Ziniao did not say) counts as agreement: the check exists to
    catch a contradiction, not to fail stores whose listing lacks the field.
    """
    if not ziniao_platform_name:
        return True
    return bool(platform(platform_code).ziniao_pattern.search(str(ziniao_platform_name)))


def platform_from_ziniao(ziniao_platform_name):
    """Best guess of our platform code from Ziniao's name, or None. Used only to
    suggest config entries (list_stores), never to decide what runs."""
    if ziniao_platform_name and GENERIC_ZINIAO.search(str(ziniao_platform_name).strip()):
        return None
    for p in PLATFORMS.values():
        if ziniao_platform_name and p.ziniao_pattern.search(str(ziniao_platform_name)):
            return p.code
    return None


# ------------------------------------------------------------------ run units
#
# The runner, the window and the run log call one market of a run a "unit". For
# MercadoLibre the unit is the bare country ("MX", "CL", "PE") — exactly the
# "region" those places used before other platforms existed, so MercadoLibre's
# commands, log tags ("[CL] "), alert files and task arguments stay character for
# character the same. Other platforms' units are their market ids ("FA-CL").

def unit_of_market(market):
    p, c = split(market)
    return c if p == DEFAULT_PLATFORM else market_id(p, c)


def unit_market(unit):
    """'CL' -> 'ML-CL'; 'FA-CL' -> 'FA-CL' (validated)."""
    return normalize(unit)


def unit_platform(unit):
    return split(unit)[0]


def unit_label(unit):
    """'CL' -> '智利' (as before); 'FA-CL' -> 'Falabella 智利'."""
    p, c = split(unit)
    return country_name(c) if p == DEFAULT_PLATFORM and SEP not in str(unit) else label(unit)


def all_units():
    """Every unit the code supports: MercadoLibre's countries first, then the others."""
    return [unit_of_market(m) for m in all_markets()]
