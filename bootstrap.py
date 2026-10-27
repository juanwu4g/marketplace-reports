# -*- coding: utf-8 -*-
"""
Environment check / setup for the ziniao automation on a new machine.

    python bootstrap.py              # check everything, report what's missing
    python bootstrap.py --drivers    # also download the fallback chromedriver set (~300 MB)

Checks, in order:
  1. Python version (and whether you're inside .venv)
  2. every package of requirements.txt installed, at the pinned version
  3. config.json present and complete
  4. the store setup: allowlist, catalog, batches, every platform package and its
     country facts (reads config.json only; opens nothing), and the optional
     SKU cost table (reports.sku_costs)
  5. ziniao client located on disk
  6. ELECTRON_RUN_AS_NODE not poisoning the environment
  7. fallback chromedriver folder (optional)

Exits non-zero if anything required is missing.
"""
import argparse
import hashlib
import io
import json
import os
from core import console  # noqa: F401  中文输出编码保护
import platform
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT, "config.json")

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"

OK, BAD, WARN = "[ 通过 ]", "[ 失败 ]", "[ 警告 ]"
problems = []


def fail(msg, fix):
    problems.append((msg, fix))
    print("{} {}".format(BAD, msg))
    print("       解决：{}".format(fix))


# --- candidate install locations, checked in order -------------------------
WIN_CLIENT_CANDIDATES = [
    r"D:\ziniao\ziniao.exe",
    r"C:\ziniao\ziniao.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\ziniao\ziniao.exe"),
    os.path.expandvars(r"%PROGRAMFILES%\ziniao\ziniao.exe"),
    # V5 client is named starter.exe
    r"D:\SuperBrowser\starter.exe",
    r"C:\SuperBrowser\starter.exe",
]
MAC_CLIENT_CANDIDATES = ["/Applications/ziniao.app"]
LINUX_CLIENT_CANDIDATES = ["/opt/ziniao/ziniaobrowser"]


def find_client():
    if IS_WIN:
        cands = WIN_CLIENT_CANDIDATES
    elif IS_MAC:
        cands = MAC_CLIENT_CANDIDATES
    else:
        cands = LINUX_CLIENT_CANDIDATES
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return None
    with io.open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def check_python():
    v = sys.version_info
    mark = OK if v >= (3, 8) else BAD
    print("{} python {}.{}.{}".format(mark, v.major, v.minor, v.micro))
    print("       {}".format(sys.executable))
    if v < (3, 8):
        fail("需要 python 3.8 或更高版本", "安装更新的 python")
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    if not in_venv:
        print("{} 当前不在 .venv 环境中，依赖可能被装到全局".format(WARN))
        print("       解决：python -m venv .venv，之后使用 venv 里的 python")


def requirements():
    """[(distribution, pinned version)] from requirements.txt."""
    out = []
    try:
        with io.open(os.path.join(ROOT, "requirements.txt"), encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if "==" in line:
                    name, ver = line.split("==", 1)
                    out.append((name.strip(), ver.strip()))
    except (IOError, OSError):
        fail("找不到 requirements.txt", "确认在项目文件夹里运行")
    return out


def check_deps():
    try:
        from importlib import metadata
    except ImportError:                                  # python < 3.8
        metadata = None
    for name, pin in requirements():
        try:
            ver = metadata.version(name) if metadata else "?"
        except Exception:
            fail("{} 未安装（要求 {}）".format(name, pin),
                 "先 set PYTHONUTF8=1，再 .venv\\Scripts\\pip install -r requirements.txt")
            continue
        if ver == pin:
            print("{} {} {}".format(OK, name, ver))
        else:
            print("{} {} {}  （要求版本：{}）".format(WARN, name, ver, pin))


def check_stores(cfg):
    """The store setup as the runner will read it. Every problem the runner would
    stop on is found here first, without opening Ziniao or any store."""
    if cfg is None:
        return
    from core import markets
    from core import store_config
    try:
        allowed = store_config.allowed_stores()
    except store_config.StoreConfigError as e:
        fail("店铺配置读不了：{}".format(e), "修正 config.json 的 stores 部分")
        return
    if not allowed:
        fail("stores.allowed 是空的", "在 config.json 的 stores.allowed 里列出可以打开的店铺")
        return
    # 条目写错 / 平台这一版程序不认识 / 没登记的店：只警告，别的店照常检查。
    # 以前一家写错整份 config 判"失败"，加一家新平台的店就把所有店拦住了。
    for name, why in store_config.config_problems():
        print("{} 店铺 {} 这一版程序跑不了：{}".format(WARN, name, why))
        print("       它不会参与任何一轮；别的店不受影响。修正 config.json 的 stores.catalog")
    try:
        found = store_config.markets_configured()
    except store_config.StoreConfigError as e:
        fail("店铺配置有误：{}".format(e), "修正 config.json 的 stores.catalog / batches")
        return
    import importlib
    for market in found:
        code, country = markets.split(market)
        try:
            names = store_config.stores_for_market(market)
        except store_config.StoreConfigError as e:
            fail("{}：{}".format(markets.label(market), e), "修正 config.json 的 stores.batches")
            continue
        pkg = markets.platform(code).package
        try:
            importlib.import_module(pkg + ".run_batch")
            importlib.import_module(pkg + ".run_reports")
        except Exception as e:
            fail("{} 的程序包 {} 导入失败：{}".format(markets.label(market), pkg, str(e)[:120]),
                 "代码不完整：重新 git pull，或运行 update.cmd")
            continue
        try:
            site = importlib.import_module(pkg + ".site")
        except ImportError:
            site = None
        if site is not None and hasattr(site, "get"):
            try:
                site.get(country)
            except Exception as e:
                fail("{}：{}".format(markets.label(market), e), "这个国家还没有站点资料")
                continue
        print("{} {}：{} 家店（{}）".format(OK, markets.label(market), len(names),
                                         "、".join(names[:4]) + ("…" if len(names) > 4 else "")))


def check_config():
    try:
        cfg = load_config()
    except ValueError as e:
        fail("config.json 不是合法 JSON：{}".format(e), "修正文件语法")
        return None
    if cfg is None:
        fail("缺少 config.json",
             "把 config.example.json 复制为 config.json 并填写账号信息")
        return None

    ui = cfg.get("user_info") or {}
    missing = [k for k in ("company", "username", "password")
               if not ui.get(k) or str(ui[k]).startswith("YOUR_")]
    if missing:
        fail("config.json 的 user_info 不完整：{}".format(", ".join(missing)),
             "填写企业登录的公司名/用户名/密码")
    else:
        print("{} config.json  company={}  port={}".format(
            OK, ui.get("company"), cfg.get("socket_port")))
    return cfg


def check_sku_costs(cfg):
    """The optional SKU cost table (config.json reports.sku_costs): configured
    but unusable would put an error on every workbook — say so before a run."""
    if cfg is None:
        return
    from core import sku_costs
    try:
        path = sku_costs.configured_path(CONFIG_PATH)
    except Exception as e:
        fail("成本表配置读不出来：{}".format(e), "修正 config.json 的 reports.sku_costs")
        return
    if not path:
        print("{} 未配置 SKU 成本表（reports.sku_costs）：报表不算商品成本".format(OK))
        return
    t = sku_costs.load(path)
    if not t.usable:
        fail(t.error, "把成本表放到 config.json 写的位置，或改 reports.sku_costs")
        return
    print("{} SKU 成本表 {}：{} 行可用{}".format(
        OK, t.name, len(t.rows),
        "，{} 行有问题（报表 ⑩ 页会列出）".format(len(t.problems)) if t.problems else ""))


def check_client(cfg):
    path = (cfg or {}).get("client_path")
    if path and os.path.exists(path):
        print("{} 紫鸟客户端（来自配置）：{}".format(OK, path))
        return path
    if path:
        print("{} 配置里的 client_path 不存在：{}".format(WARN, path))
    found = find_client()
    if found:
        print("{} 紫鸟客户端（自动检测）：{}".format(OK, found))
        print('       如需固定路径，可在 config.json 里设置 "client_path"')
        return found
    fail("未找到紫鸟客户端",
         "安装紫鸟浏览器，或在 config.json 里设置 client_path")
    return None


def check_electron_env():
    if os.environ.get("ELECTRON_RUN_AS_NODE"):
        print("{} 当前环境设置了 ELECTRON_RUN_AS_NODE".format(WARN))
        print("       启动器会在子进程里清掉它，所以本身没问题，")
        print("       但不要用裸 os.system() 去启动客户端。")
    else:
        print("{} 未设置 ELECTRON_RUN_AS_NODE".format(OK))


def driver_dir(cfg):
    d = (cfg or {}).get("driver_folder_path") or "webdriver"
    return d if os.path.isabs(d) else os.path.join(ROOT, d)


def check_drivers(cfg, download):
    d = driver_dir(cfg)
    if os.path.isdir(d):
        present = sorted(f for f in os.listdir(d) if f.startswith("chromedriver"))
    else:
        present = []
    print("{} 备用 chromedriver：{} 个，位于 {}".format(
        OK if present else WARN, len(present), d))
    if not present:
        print("       通常无需处理：每个店铺内核自带 webdriver.exe，")
        print("       优先使用它。需要备用驱动时加 --drivers 下载。")
    if download:
        download_drivers(d)


DRIVER_CONFIG_URLS = {
    "win": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/exe_32/config.json",
    "mac_x64": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/mac/x64/config.json",
    "mac_arm": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/mac/arm64/config.json",
}


def download_drivers(dest):
    import requests
    if IS_WIN:
        url = DRIVER_CONFIG_URLS["win"]
    elif IS_MAC:
        key = "mac_arm" if platform.machine() == "arm64" else "mac_x64"
        url = DRIVER_CONFIG_URLS[key]
    else:
        print("       Linux 使用内核自带驱动，无需下载")
        return
    print("--- 正在下载 chromedriver 到 {} ---".format(dest))
    cfg = json.loads(requests.get(url, timeout=60).text)
    if not os.path.isdir(dest):
        os.makedirs(dest)
    for item in cfg:
        name = item["name"] + (".exe" if IS_WIN else "")
        path = os.path.join(dest, name)
        if os.path.exists(path):
            with open(path, "rb") as f:
                if hashlib.new("sha1", f.read()).hexdigest() == item["sha1"]:
                    print("  已存在   {}".format(name))
                    continue
        print("  下载中   {}".format(name))
        r = requests.get(item["url"], stream=True, timeout=300)
        if r.status_code != 200:
            print("           失败，状态码 {}".format(r.status_code))
            continue
        with open(path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        if not IS_WIN:
            os.chmod(path, 0o755)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--drivers", action="store_true",
                    help="download the fallback chromedriver set (~300 MB)")
    args = ap.parse_args()

    from core import runlog
    print("=" * 62)
    print(" 紫鸟自动化 —— 环境自检")
    print(" 代码版本：{}".format(runlog.code_version(ROOT)))
    print(" 系统平台：{} {}".format(platform.system(), platform.machine()))
    print("=" * 62)
    check_python()
    check_deps()
    cfg = check_config()
    check_stores(cfg)
    check_sku_costs(cfg)
    check_client(cfg)
    check_electron_env()
    check_drivers(cfg, args.drivers)
    print("=" * 62)
    if problems:
        print(" 运行前必须先解决 {} 个问题：".format(len(problems)))
        for msg, fix in problems:
            print("   - {}".format(msg))
            print("       -> {}".format(fix))
        return 1
    print(" 环境检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
