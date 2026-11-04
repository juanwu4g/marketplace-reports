# -*- coding: utf-8 -*-
"""
Single source of truth for which store the tooling touches.

The store name used to be hardcoded in six places (`run_downloads.py`,
`open_store.py`, both analysis scripts and twice in `meli_data.py`), which meant
there was no way to see - or limit - what the tooling could open. It now lives
in `config.json`:

    "stores": {
      "default": "BOCINA_SM",
      "allowed": ["BOCINA_SM", "EWTTO_SM"],
      "batch": []
    }

`default`  the store used when no --store is given.
`allowed`  the ONLY stores that may be opened. An empty list means no
           restriction, which is the previous behaviour; fill it in to make the
           scope explicit and auditable.
`batch`    which stores `run_batch.py` walks, in order. Empty (or absent)
           means "every store in `allowed`", which is what you usually want -
           it exists for the case where the allowlist is deliberately wider
           than the nightly run.

Deliberately dependency-free: `meli_data.py` imports this, and it is in turn
imported by the analysis scripts, which have no business loading selenium or
requests just to learn a store name.
"""
import io
from core import console  # noqa: F401  中文店名输出保护
import json
import os

from core.paths import ROOT
CONFIG_PATH = os.path.join(ROOT, "config.json")



class StoreNotAllowed(Exception):
    """Raised before any request reaches Ziniao."""


class StoreConfigError(Exception):
    """config.json is unreadable, so the allowlist cannot be trusted."""


def _load():
    """Read config.json, or fail loudly.

    This MUST NOT swallow a parse error. An earlier version returned {} on
    malformed JSON, which made `allowed` fall back to [] - meaning
    "unrestricted" - so a single missing comma silently unlocked all 93 stores.
    A guard that fails open is not a guard. Fail closed instead.
    """
    try:
        with io.open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except ValueError as e:
        raise StoreConfigError(
            "config.json is not valid JSON (%s). Refusing to run: a malformed "
            "config cannot be trusted to limit which stores may be opened." % e)
    except (IOError, OSError) as e:
        raise StoreConfigError("cannot read %s (%s)" % (CONFIG_PATH, e))


def _stores():
    cfg = _load().get("stores")
    return cfg if isinstance(cfg, dict) else {}


def default_store():
    """The store used when nothing is specified."""
    name = _stores().get("default")
    if not name:
        # 绝不内置一个店名兜底。早先这里写死了某家店，config.json 缺少
        # stores.default 时会静默打开它 —— 白名单为空（不限制）的情况下
        # 就是在没人要求的店上跑了一整轮。缺配置要报错，不要猜。
        raise StoreConfigError(
            "config.json 里没有 stores.default。请在其中指定默认店铺，"
            "或在命令行用 --store / --stores 明确指定。")
    return name


def allowed_stores():
    """Names that may be opened, or [] meaning unrestricted."""
    allowed = _stores().get("allowed")
    return list(allowed) if isinstance(allowed, list) else []


def batch_stores():
    """Stores run_batch.py should walk, in order.

    Defaults to the whole allowlist. Returns [] when the allowlist is empty -
    i.e. unrestricted - because "walk every store on the account" is never an
    intended batch; callers must treat [] as "nothing configured" and refuse
    rather than iterating 93 stores.
    """
    batch = _stores().get("batch")
    if isinstance(batch, list) and batch:
        return [require_allowed(n) for n in batch]
    if _uses_catalog():
        # 新写法没有 "batch" 这个键：美客多不带店铺跑的那一批是墨西哥那一批
        # （batches["ML-MX"]，没写就是白名单里的美客多墨西哥店）。不能退回"白名单里
        # 全部美客多的店" —— 那样墨西哥那一轮会把智利、秘鲁的店也拿去墨西哥站点跑。
        return stores_for_market(markets.market_id(markets.DEFAULT_PLATFORM,
                                                   markets.DEFAULT_COUNTRY))
    # 旧配置（没有 catalog）：和以前一样是整份白名单。
    return [n for n in _resolvable(allowed_stores())
            if platform_of(n) == markets.DEFAULT_PLATFORM]


def _resolvable(names):
    """names 里能定出市场的店（config 写错的那几家跳过：config_problems 会说出来）。"""
    out = []
    for n in names:
        try:
            market_of(n)
        except StoreConfigError:
            continue
        out.append(n)
    return out


def is_allowed(name):
    allowed = allowed_stores()
    return True if not allowed else name in allowed


def require_allowed(name):
    """Raise unless `name` may be opened. Call before touching the client.

    Failing here rather than at startBrowser means an unintended store is
    rejected locally, without a request ever reaching Ziniao.
    """
    if is_allowed(name):
        return name
    raise StoreNotAllowed(
        "store %r is not in config.json -> stores.allowed (%s). "
        "Add it there if you meant to open it."
        % (name, ", ".join(allowed_stores()) or "empty"))


def describe():
    allowed = allowed_stores()
    return "default=%s allowed=%s" % (
        _stores().get("default") or "(未配置)",
        ", ".join(allowed) if allowed else "(不限制)")


# ---------------------------------------------------------------- 市场（平台 × 国家）
#
# 一家店属于一个**市场** = 一个平台的一个国家："ML-MX"（美客多墨西哥）、
# "FA-CL"（Falabella 智利）。平台和国家的清单在 core/markets.py。
#
# 推荐的写法（新）：
#
#     "stores": {
#       "allowed": [... 能打开的全部店铺 ...],
#       "catalog": {
#         "BOCINA_SM":          {"platform": "ML", "country": "MX"},
#         "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"}
#       },
#       "batches": {"ML-MX": [...], "ML-CL": [...], "FA-CL": [...]}
#     }
#
# 店铺在紫鸟里**改了名**（例如 清货2_EE02 改成 QH_EE02）：白名单、catalog、batches 里
# 一律写新名，并在 catalog 这一家的条目里用 "formerly" 写上旧名（可以写几个，
# 也可以只写一个字符串）。硬盘上什么都不用搬：旧名下的月度结果和下载目录会
# 自动算作这家店的，季度报表、本月总览只显示新名、只算一次。
#
#       "catalog": {
#         "QH_EE02": {"platform": "ML", "country": "MX", "formerly": ["清货2_EE02"]},
#         "DAB_CL Admin(Coco)": {"platform": "FA", "country": "CL",
#                                "formerly": "DAB_CL (Coco) Admin"}
#       }
#
# 规则：旧名不能同时还是一家现在的店（白名单或 catalog 里有它），也不能被两家店
# 同时认领 —— 写错的那家店算配置有问题（config_problems 说出来），别的店照常。
# "formerly" 只认 catalog 这种写法；旧写法（regions）没有地方写它。
#
# 旧写法照样认（生产机器上的 config.json 就是这么写的，一个字不用改）：
#
#     "batch":   [... 美客多墨西哥那一批 ...]          = batches["ML-MX"]
#     "regions": {"Tecnohub_CL": "CL"}                 = catalog 里 platform 为 ML
#     "batches": {"CL": [...]}                         = batches["ML-CL"]
#
# 规则：
# * catalog 优先于 regions。
# * **只要写了 catalog，白名单里每一家店都必须在 catalog（或 regions）里写明**
#   属于哪个市场，否则报错。没写 catalog 的旧配置，没写明的店算美客多墨西哥
#   —— 第二个平台出现之前每家店都是墨西哥店。
# * 平台、国家代码不认识，或这个平台还不支持那个国家 -> 报错（fail closed）。
# * 一批店里混进别的市场的店 -> 报错。一家 Falabella 店抄进美客多智利那一批，
#   会被拿去跑美客多的页面、汇进美客多的全店汇总 —— 宁可报错。
#
# 这里只认 config.json，不看店名、也不看紫鸟的平台名：决定"这一轮跑哪些店、
# 报表叫什么名字"的规则必须写在配置里、人能核对。紫鸟的平台名只用来**反查**
# （check_ziniao_platform）：配置说是 Falabella、紫鸟说是 MercadoLibre，就拒绝。

from core import markets  # noqa: E402

DEFAULT_REGION = markets.DEFAULT_COUNTRY
# 美客多支持的国家（窗口里"选国家"、run_reports --region 用）
REGION_NAMES = dict((c, markets.country_name(c))
                    for c in markets.platform("ML").countries)


def region_name(region):
    """界面和日志里用的中文名：MX -> 墨西哥；FA-CL -> Falabella 智利。不认识的代码原样返回。"""
    if markets.SEP in str(region or ""):
        try:
            return markets.label(markets.normalize(region))
        except markets.UnknownMarket:
            return str(region)
    return markets.country_name(region)


def _catalog_scan():
    """config 里写明了市场的店，和写得有问题的店。

    返回 ({店名: (平台, 国家)}, {店名: 一句原因})。一家店的条目写错（平台这一版程序
    不认识、国家不支持、格式不对）只影响这家店：它进"问题"那一份，别的店照常。以前
    一条写错整份 config 都拒收，于是别的平台、别的国家那一个月全都不跑，窗口里的
    店铺树也是空的 —— 加一家还没合并进来的新平台的店就能把整台机器停掉。
    """
    cfg = _stores()
    out, bad = {}, {}
    regions = cfg.get("regions")
    if isinstance(regions, dict):
        for name, country in regions.items():
            if not country:
                continue
            _place(out, bad, name, markets.DEFAULT_PLATFORM, country, "stores.regions")
    cat = cfg.get("catalog")
    if isinstance(cat, dict):
        for name, entry in cat.items():
            if not isinstance(entry, dict):
                bad[name] = ('stores.catalog 里的写法不对，应该写成 '
                             '{"platform": "ML", "country": "MX"}')
                out.pop(name, None)
                continue
            _place(out, bad, name, entry.get("platform"), entry.get("country"),
                   "stores.catalog")
        _scan_formerly(cat, out, bad)
    return out, bad


def _formerly_of(entry):
    """The "formerly" field of a catalog entry as a list of names. Raises
    ValueError when it is neither a string nor a list of strings."""
    raw = entry.get("formerly")
    if raw in (None, ""):
        return []
    if isinstance(raw, str):
        raw = [raw]
    if not isinstance(raw, list) or not all(isinstance(x, str) for x in raw):
        raise ValueError("formerly")
    out = []
    for x in raw:
        x = x.strip()
        if x and x not in out:
            out.append(x)
    return out


def _scan_formerly(cat, out, bad):
    """Check every valid catalog entry's "formerly" names. A problem is the
    claiming store's (it goes to `bad` and out of `out`): a former name that is
    still a current store (allowlist, catalog or regions), or one claimed by two
    stores. Never a crash — the other stores run as usual."""
    current = set(out) | set(bad) | set(allowed_stores())
    claims = {}                                  # former name -> [claiming stores]
    for name, entry in cat.items():
        if name in bad or not isinstance(entry, dict):
            continue
        try:
            olds = _formerly_of(entry)
        except ValueError:
            bad[name] = ('stores.catalog 里 %s 的 formerly 写法不对，应该写成 ["旧店名"]' % name)
            continue
        for old in olds:
            if old == name:
                bad[name] = "stores.catalog 里 %s 的 formerly 写了它自己" % name
                break
            if old in current:
                bad[name] = ("stores.catalog 里 %s 的 formerly 写了 %s，但 %s 还是一家现在的店"
                             "（白名单或 catalog 里有它）：旧名不能同时是现在的店名" % (name, old, old))
                break
            claims.setdefault(old, []).append(name)
    for old, names in claims.items():
        if len(names) > 1:
            for name in names:
                bad[name] = ("旧名 %s 同时写在 %s 的 formerly 里，一个旧名只能属于一家店"
                             % (old, "、".join(names)))
    for name in list(bad):
        out.pop(name, None)


def _place(out, bad, name, platform, country, where):
    try:
        out[name] = markets.check(platform, country)
        bad.pop(name, None)
    except markets.UnknownMarket as e:
        bad[name] = "%s 里的条目：%s" % (where, e)
        out.pop(name, None)


def _catalog_entries():
    """config 里写明了市场的店：{店名: (平台, 国家)}，已校验（写错的不在里面）。"""
    return _catalog_scan()[0]


def config_problems():
    """config.json 里这一版程序跑不了的店：[(店名, 原因)]，按 config 的顺序。

    写错的 catalog 条目，和白名单里有、catalog 里没写的店（用了 catalog 的配置）。
    运行器、窗口、bootstrap 都用它把这些店说出来，而不是整份 config 拒收。
    """
    entries, bad = _catalog_scan()
    out = list(bad.items())
    if _uses_catalog():
        for name in allowed_stores():
            if name not in entries and name not in bad:
                out.append((name, "白名单里有，stores.catalog 里没写它属于哪个平台和国家"))
    return out


def problem_stores():
    return set(n for n, _ in config_problems())


# ---------------------------------------------------------------- 改过名的店（formerly）

def former_map():
    """{旧名: 现在的店名}，只取 catalog 里写法正确的店（写错的店不认领任何旧名）。"""
    entries, _bad = _catalog_scan()
    cat = _stores().get("catalog")
    out = {}
    if not isinstance(cat, dict):
        return out
    for name in entries:
        entry = cat.get(name)
        if not isinstance(entry, dict):
            continue
        try:
            for old in _formerly_of(entry):
                out[old] = name
        except ValueError:
            continue
    return out


def former_names(store):
    """这家店在紫鸟里以前叫过的名字（catalog 条目的 "formerly"），没有就是 []。
    条目写错的店返回 []：它不认领任何旧名。"""
    return [old for old, cur in former_map().items() if cur == store]


def current_name(name):
    """`name` 是某家店的旧名就返回那家店现在的名字，否则原样返回。"""
    return former_map().get(name, name)


def catalog_option(store, key, default=None):
    """An extra field of a store's stores.catalog entry (beyond platform and
    country), e.g. Liverpool's "marketplace_store": the second Ziniao store a
    platform needs for the same seller. None / default when not set."""
    cat = _stores().get("catalog")
    entry = cat.get(store) if isinstance(cat, dict) else None
    if not isinstance(entry, dict):
        return default
    value = entry.get(key)
    return default if value in (None, "") else value


def companion_store(store, key):
    """The second Ziniao store named by catalog_option(store, key), checked like
    a store before Ziniao is touched: it must be on the allowlist and must not be
    a store of another platform. None when the option is not set."""
    name = catalog_option(store, key)
    if name is None:
        return None
    name = require_allowed(str(name))
    known = _catalog_entries()
    if name in known and known[name][0] != platform_of(store):
        raise StoreConfigError(
            "config.json stores.catalog[%r].%s = %r，但 %r 是 %s 的店"
            % (store, key, name, name, markets.market_id(*known[name])))
    return name


def _uses_catalog():
    cat = _stores().get("catalog")
    return isinstance(cat, dict) and bool(cat)


def market_parts(store):
    """(平台, 国家)。没写明的店：旧配置算 ML-MX；用了 catalog 的配置报错。
    条目写错的店报错，说出原因（别的店不受影响）。"""
    known, bad = _catalog_scan()
    if store in known:
        return known[store]
    if store in bad:
        raise StoreConfigError("config.json 里的 %s：%s" % (store, bad[store]))
    if _uses_catalog():
        raise StoreConfigError(
            "config.json stores.catalog 里没有 %r：请写明它属于哪个平台和国家，"
            '例如 "%s": {"platform": "ML", "country": "MX"}' % (store, store))
    return markets.DEFAULT_PLATFORM, markets.DEFAULT_COUNTRY


def market_of(store):
    """'FA-CL'。"""
    return markets.market_id(*market_parts(store))


def platform_of(store):
    return market_parts(store)[0]


def country_of(store):
    return market_parts(store)[1]


def region_of(store):
    """这家店属于哪个国家（"MX" / "CL"）。旧名字，等同 country_of。"""
    return country_of(store)


def configured_country(store):
    """config 里明确写了的国家，没写返回 None（站点判断时再按别的线索猜）。"""
    entry = _catalog_entries().get(store)
    return entry[1] if entry else None


def _listed_batch(market):
    """batches[market]，以及旧写法的等价键。没写返回 None。"""
    cfg = _stores()
    batches = cfg.get("batches") if isinstance(cfg.get("batches"), dict) else {}
    upper = dict((str(k).upper(), v) for k, v in batches.items())
    p, c = markets.split(market)
    candidates = [market]
    where = ["batches." + market]
    if p == markets.DEFAULT_PLATFORM:
        candidates.append(c)
        where.append("batches." + c)
    for key, label in zip(candidates, where):
        v = upper.get(key)
        if isinstance(v, list) and v:
            return v, label
    if market == markets.market_id(markets.DEFAULT_PLATFORM, markets.DEFAULT_COUNTRY):
        v = cfg.get("batch")
        if isinstance(v, list) and v:
            return v, "batch"
    return None, None


def stores_for_market(market):
    """某个市场"全部店铺"指哪几家，按顺序。

    写了这个市场的批次（batches["FA-CL"]，或旧写法）就用它，每家店都要过白名单
    并且必须真的属于这个市场；没写就取白名单里属于这个市场的店。
    """
    market = markets.normalize(market)
    listed, where = _listed_batch(market)
    if listed:
        names = [require_allowed(n) for n in listed]
        wrong = [n for n in names if market_of(n) != market]
        if wrong:
            raise StoreConfigError(
                "config.json 里 %s（%s）那一批混进了别的市场的店：%s。"
                "检查 stores.catalog / stores.regions 和 stores.%s。"
                % (market, markets.label(market), "、".join(
                    "%s（%s）" % (n, market_of(n)) for n in wrong), where))
        return names
    return [n for n in _resolvable(allowed_stores()) if market_of(n) == market]


def stores_for_region(region):
    """美客多某个国家的全部店铺（旧名字）= stores_for_market("ML-<国家>")。

    MX：沿用 `batch`；`batch` 为空时取白名单里属于美客多墨西哥的店。别的平台的
    店永远不会出现在这里。
    """
    if markets.SEP in str(region or ""):          # 别的平台的运行单元："FA-CL"
        return stores_for_market(region)
    return stores_for_market(markets.market_id(markets.DEFAULT_PLATFORM,
                                               region or DEFAULT_REGION))


def markets_configured():
    """有店铺的市场，美客多墨西哥排第一，其余按白名单里第一次出现的顺序。"""
    cfg = _stores()
    found = [market_of(n) for n in _resolvable(allowed_stores())]
    found += [markets.market_id(*v) for v in _catalog_entries().values()]
    batches = cfg.get("batches")
    if isinstance(batches, dict):
        for k, v in batches.items():
            if v:
                try:
                    found.append(markets.normalize(k))
                except markets.UnknownMarket:
                    continue             # 这一版程序不认识的市场：config_problems 说出来
    first = markets.market_id(markets.DEFAULT_PLATFORM, markets.DEFAULT_COUNTRY)
    out = [first] if (first in found or not found) else []
    for m in found:
        if m not in out:
            out.append(m)
    return out


def regions():
    """美客多有店铺的国家，墨西哥排第一。白名单为空（不限制）时只有墨西哥。"""
    out = [markets.split(m)[1] for m in markets_configured()
           if markets.split(m)[0] == markets.DEFAULT_PLATFORM]
    return out or [DEFAULT_REGION]


def group_by_market(stores):
    """[店名] -> [(市场, [店名])]，保持原顺序。"""
    out = []
    for s in stores:
        m = market_of(s)
        for key, names in out:
            if key == m:
                names.append(s)
                break
        else:
            out.append((m, [s]))
    return out


def group_by_region(stores):
    """[店名] -> [(国家, [店名])]，保持原顺序。界面重跑时用：一轮只跑一个区域。"""
    out = []
    for s in stores:
        r = region_of(s)
        for reg, names in out:
            if reg == r:
                names.append(s)
                break
        else:
            out.append((r, [s]))
    return out


def require_platform(names, platform):
    """这一批店必须全是 `platform` 的店，否则报错（在碰紫鸟之前）。"""
    platform = markets.platform(platform).code
    wrong = [n for n in names if platform_of(n) != platform]
    if wrong:
        raise StoreConfigError(
            "%s 的下载/报表不能跑别的平台的店：%s。检查 config.json stores.catalog。"
            % (markets.platform(platform).name_zh, "、".join(
                "%s（%s）" % (n, market_of(n)) for n in wrong)))
    return list(names)


def check_ziniao_platform(store, ziniao_platform_name, platform=None):
    """开店前反查：紫鸟说的平台名和 config 写的平台对不上就报错。

    紫鸟没给平台名时放行（这道检查只抓矛盾）。
    """
    platform = platform or platform_of(store)
    if not markets.ziniao_matches(platform, ziniao_platform_name):
        raise StoreConfigError(
            "店铺 %r 在 config.json 里是 %s，紫鸟却说它是 %r 平台的店。"
            "不打开它：检查 stores.catalog 是不是写错了。"
            % (store, markets.platform(platform).name_zh, ziniao_platform_name))
    return True


# ---------------------------------------------------------------- 运行单元
#
# 运行器、窗口和运行日志按"单元"跑：美客多的单元就是国家代码（MX / CL / PE，和
# 以前的 region 一字不差），别的平台的单元是市场代码（FA-CL）。见 markets.py。

def unit_of(store):
    """这家店属于哪个运行单元：美客多的店 -> "CL"；Falabella 的店 -> "FA-CL"。"""
    return markets.unit_of_market(market_of(store))


def other_units_configured():
    """美客多以外、config 里有店铺的运行单元，按白名单顺序。"""
    return [markets.unit_of_market(m) for m in markets_configured()
            if markets.split(m)[0] != markets.DEFAULT_PLATFORM]

