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
    return allowed_stores()


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


# ---------------------------------------------------------------- 区域（国家）
#
# 同一个紫鸟账号下既有墨西哥店也有智利店。config.json 里：
#
#     "stores": {
#       "allowed": [... 两国全部店铺 ...],
#       "batch":   [... 墨西哥那一批，顺序即下载顺序 ...],
#       "regions": {"Tecnohub_CL": "CL", ...},
#       "batches": {"CL": [... 智利那一批 ...]}
#     }
#
# 没写进 regions 的店一律算墨西哥 —— 区域出现之前每家店都是墨西哥店，旧配置
# 一个字不改照样能跑。`batch` 也因此还是"墨西哥那一批"，不改名：生产机器上
# 的 config.json 就是这么写的。
#
# 这里只认 config.json，不看店名里的 _CL、也不看紫鸟的平台名（meli_site 会
# 看）：决定"这一轮跑哪些店、报表叫什么名字"的规则必须是写在配置里、人能
# 核对的，不能靠猜。

DEFAULT_REGION = "MX"
REGION_NAMES = {"MX": "墨西哥", "CL": "智利", "PE": "秘鲁"}


def region_name(region):
    """界面和日志里用的中文名：MX -> 墨西哥。不认识的代码原样返回。"""
    return REGION_NAMES.get(str(region or "").upper(), str(region or ""))


def _region_map():
    regions = _stores().get("regions")
    if not isinstance(regions, dict):
        return {}
    return dict((k, str(v).upper()) for k, v in regions.items() if v)


def region_of(store):
    """这家店属于哪个区域（"MX" / "CL"）。config 里没写就是墨西哥。"""
    return _region_map().get(store, DEFAULT_REGION)


def stores_for_region(region):
    """某个区域"全部店铺"指哪几家，按顺序。

    MX：沿用 `batch`（以前的"全部店铺"就是它，一字不变）；`batch` 为空时取
    白名单里属于墨西哥的店。注意不是 batch_stores() —— 它在 batch 为空时会把
    整份白名单（含智利店）都返回。
    其他区域：`batches[区域]`；没写就取白名单里标成这个区域的店。

    每家店都要过白名单，并且必须真的属于这个区域：一家智利店误抄进墨西哥的
    batch，会被拿去墨西哥站点下载、汇进墨西哥的全店汇总 —— 宁可报错。
    """
    region = str(region or DEFAULT_REGION).upper()
    cfg = _stores()
    if region == DEFAULT_REGION:
        listed = cfg.get("batch")
    else:
        listed = (cfg.get("batches") or {}).get(region) \
            if isinstance(cfg.get("batches"), dict) else None
    if isinstance(listed, list) and listed:
        names = [require_allowed(n) for n in listed]
        wrong = [n for n in names if region_of(n) != region]
        if wrong:
            raise StoreConfigError(
                "config.json 里 %s 那一批混进了别的区域的店：%s。"
                "检查 stores.regions 和 stores.%s。"
                % (region, "、".join("%s（%s）" % (n, region_of(n)) for n in wrong),
                   "batch" if region == DEFAULT_REGION else "batches." + region))
        return names
    return [n for n in allowed_stores() if region_of(n) == region]


def regions():
    """有店铺的区域，墨西哥排第一。白名单为空（不限制）时只有墨西哥。"""
    cfg = _stores()
    found = [region_of(n) for n in allowed_stores()]
    found += list(_region_map().values())
    if isinstance(cfg.get("batches"), dict):
        found += [str(k).upper() for k, v in cfg["batches"].items() if v]
    out = [DEFAULT_REGION] if (DEFAULT_REGION in found or not found) else []
    for r in found:
        if r not in out:
            out.append(r)
    return out


def group_by_region(stores):
    """[店名] -> [(区域, [店名])]，保持原顺序。界面重跑时用：一轮只跑一个区域。"""
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
