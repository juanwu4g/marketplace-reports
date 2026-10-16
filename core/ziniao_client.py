# -*- coding: utf-8 -*-
"""
Ziniao (紫鸟) SuperBrowser client wrapper.

Handles the parts the vendor demo gets wrong or leaves implicit:
  - strips ELECTRON_RUN_AS_NODE when spawning the client (see GUIDE.md 6.1)
  - caps the updateCore retry loop instead of spinning forever (6.2)
  - always leaves the browser on a normal page, never the extension IP-check
    page, so a later run can reattach (6.3)

Typical use:

    from core.ziniao_client import ZiniaoClient

    with ZiniaoClient() as zn:
        zn.start()                              # launch + wait for IPC
        store = zn.find_store("BOCINA_SM")
        with zn.open_store(store) as session:
            session.driver.get(...)
"""
import io
import json
import os
import platform
import subprocess
import time
import uuid
from contextlib import contextmanager

import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

from core import store_config

from core.paths import ROOT

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"


class ZiniaoError(RuntimeError):
    pass


# re-exported so callers can catch scope errors without importing store_config
StoreNotAllowed = store_config.StoreNotAllowed


def load_config(path=None):
    path = path or os.path.join(ROOT, "config.json")
    if not os.path.exists(path):
        raise ZiniaoError(
            "config.json not found. Copy config.example.json and fill it in.")
    with io.open(path, encoding="utf-8") as f:
        return json.load(f)


def _find_client():
    """Best-effort auto-detect, mirroring bootstrap.py."""
    if IS_WIN:
        cands = [r"D:\ziniao\ziniao.exe", r"C:\ziniao\ziniao.exe",
                 os.path.expandvars(r"%LOCALAPPDATA%\Programs\ziniao\ziniao.exe"),
                 os.path.expandvars(r"%PROGRAMFILES%\ziniao\ziniao.exe")]
    elif IS_MAC:
        cands = ["/Applications/ziniao.app"]
    else:
        cands = ["/opt/ziniao/ziniaobrowser"]
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


class StoreSession(object):
    """One opened store: the IPC response plus an attached selenium driver."""

    def __init__(self, client, info, driver):
        self.client = client
        self.info = info
        self.driver = driver

    @property
    def oauth(self):
        return self.info.get("browserOauth") or self.info.get("browserId")

    @property
    def download_path(self):
        return self.info.get("downloadPath")

    @property
    def launcher_page(self):
        return self.info.get("launcherPage")

    def check_ip(self):
        """Open the proxy check page and report whether it passed.

        Always returns to launcher_page afterwards - leaving the browser on the
        chrome-extension:// page makes future reattaches fail (GUIDE.md 6.3).
        """
        url = self.info.get("ipDetectionPage")
        if not url:
            print("  [警告] 没有 ipDetectionPage，跳过 IP 检测")
            return True
        ok = False
        try:
            self.driver.get(url)
            time.sleep(3)
            self.driver.find_element(
                By.XPATH, '//button[contains(@class, "styles_btn--success")]')
            ok = True
        except NoSuchElementException:
            print("  [警告] 未找到 IP 检测的成功标识")
        except Exception as e:
            print("  [警告] IP 检测出错：%s" % str(e)[:100])
        finally:
            if self.launcher_page:
                self.driver.get(self.launcher_page)
                time.sleep(3)
        return ok


class ZiniaoClient(object):
    def __init__(self, config=None, config_path=None):
        cfg = config or load_config(config_path)
        self.cfg = cfg
        self.user_info = cfg["user_info"]
        self.port = cfg.get("socket_port", 16851)
        self.version = cfg.get("client_version", "v6")
        self.client_path = cfg.get("client_path") or _find_client()
        if not self.client_path:
            raise ZiniaoError("ziniao client not found; set client_path in config.json")
        d = cfg.get("driver_folder_path") or "webdriver"
        self.driver_folder = d if os.path.isabs(d) else os.path.join(ROOT, d)
        self._opened = []

    # ---------------- IPC ----------------

    def send(self, data, timeout=120):
        payload = dict(data)
        payload.update(self.user_info)
        payload.setdefault("requestId", str(uuid.uuid4()))
        try:
            r = requests.post("http://127.0.0.1:%d" % self.port,
                              json.dumps(payload).encode("utf-8"), timeout=timeout)
            return json.loads(r.text)
        except Exception:
            return None

    def _require_ok(self, resp, what):
        if resp is None:
            raise self._error("%s: no response from client" % what, what, None)
        code = resp.get("statusCode")
        if code == 0:
            return resp
        text = json.dumps(resp, ensure_ascii=False)[:300]
        if code == -10003:
            raise self._error("%s: login/permission error: %s" % (what, text),
                              what, text)
        raise self._error("%s failed: %s" % (what, text), what, text)

    @staticmethod
    def _error(message, step, response):
        """ZiniaoError，外加卡在哪一步、紫鸟原话是什么。

        店铺打不开时没有浏览器可以截图，这两样就是全部的现场 —— 诊断那一步
        （download_diag）靠它们回答"为什么这家店一张报表都没有"。
        """
        e = ZiniaoError(message)
        e.ziniao_step = step
        e.ziniao_response = response or ""
        return e

    # ---------------- lifecycle ----------------

    def is_up(self):
        return self.send({"action": "getBrowserList"}, timeout=15) is not None

    def kill_existing(self):
        """Terminate a running client. This closes any stores a human has open."""
        if IS_WIN:
            name = "SuperBrowser.exe" if self.version == "v5" else "ziniao.exe"
            os.system("taskkill /f /t /im " + name + " >nul 2>&1")
        elif IS_MAC:
            os.system("killall ziniao 2>/dev/null")
        else:
            os.system("killall ziniaobrowser 2>/dev/null")
        time.sleep(4)

    def launch(self):
        args = ["--run_type=web_driver", "--ipc_type=http", "--port=%d" % self.port]
        if IS_WIN:
            cmd = [self.client_path] + args
        elif IS_MAC:
            cmd = ["open", "-a", self.client_path, "--args"] + args
        else:
            cmd = [self.client_path, "--no-sandbox"] + args

        # GUIDE.md 6.1 - inherited ELECTRON_RUN_AS_NODE makes the client exit
        # immediately with "bad option: --run_type=web_driver".
        env = os.environ.copy()
        env.pop("ELECTRON_RUN_AS_NODE", None)
        subprocess.Popen(cmd, env=env)

    def wait_ready(self, timeout=120):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.is_up():
                return True
            time.sleep(2)
        return False

    def start(self, restart=True):
        """Ensure the client is running in webdriver/http mode."""
        if self.is_up():
            print("客户端已在端口 %d 上监听" % self.port)
            return
        if restart:
            print("正在结束已运行的紫鸟客户端…")
            self.kill_existing()
        print("以 webdriver 模式启动客户端（端口 %d）…" % self.port)
        self.launch()
        if not self.wait_ready():
            raise ZiniaoError("client did not open port %d in time" % self.port)
        print("客户端就绪")

    def update_core(self, max_attempts=15):
        """Pre-download cores. Non-fatal: on some installs this never returns 0
        (GUIDE.md 6.2), and startBrowser works regardless."""
        for _ in range(max_attempts):
            r = self.send({"action": "updateCore"}, timeout=120)
            if r is None:
                time.sleep(2)
                continue
            if r.get("statusCode") == 0:
                print("内核已是最新")
                return True
            time.sleep(2)
        print("[警告] updateCore 未正常返回，继续执行")
        return False

    # ---------------- stores ----------------

    def list_stores(self):
        r = self._require_ok(self.send({"action": "getBrowserList"}), "getBrowserList")
        return r.get("browserList") or []

    def find_store(self, name=None):
        """Resolve a store name to its record.

        Every path to `startBrowser` goes through here, so this is where the
        allowlist is enforced - and it is enforced BEFORE the lookup, so an
        unintended name never reaches Ziniao at all.

        Note that `getBrowserList` returns every store the enterprise account
        can see (93 on this account) and carries no permission field, so the
        client cannot tell which are authorised for WebDriver. The allowlist in
        config.json is the only local guard against opening the wrong one.
        """
        name = name or store_config.default_store()
        store_config.require_allowed(name)
        for b in self.list_stores():
            if (b.get("browserName") or "") == name:
                return b
        raise ZiniaoError("store %r not found on this account" % name)

    DRIVER_WAIT = 600          # 内核首次下载可能要几分钟，给足
    DRIVER_SETTLE = 4          # 文件大小连续这么久不变，才认为下完了

    def _resolve_driver(self, info, wait=None):
        """找这个店铺内核对应的 webdriver。内核还在下就等它。

        首选店铺内核自带的 webdriver.exe，其次是 driver_folder 里的备用驱动。

        **必须等，不能一次查找就放弃。** 紫鸟在一台从没跑过的机器上要先下载
        店铺内核，而 startBrowser 会在内核还没落盘时就返回 —— 立刻去找
        webdriver.exe 必然找不到。老实现就是这么抛错的，结果是新机器第一次
        跑必然失败，而错误信息（"no chromedriver, run bootstrap --drivers"）
        还把人引向下载 300 MB 备用驱动这条错路。

        找到之后还要等文件大小稳定：下到一半的 webdriver.exe 比没有更糟 ——
        Selenium 会以一个看不懂的错误崩掉，而不是说"文件不完整"。
        """
        wait = self.DRIVER_WAIT if wait is None else wait
        bpath = info.get("browserPath") or ""
        if bpath.lower().endswith(("superbrowser.exe", "superbrowser")):
            bpath = os.path.dirname(bpath)
        cands = []
        if bpath:
            cands.append(os.path.join(bpath, "webdriver.exe" if IS_WIN else "webdriver"))
        major = str(info.get("core_version", "")).split(".")[0]
        if major:
            cands.append(os.path.join(
                self.driver_folder,
                "chromedriver%s%s" % (major, ".exe" if IS_WIN else "")))

        # 第一轮只看存在不存在，不做 settle 检查：调用前就已经在的文件，
        # 不可能正在下载。每家店白等 4 秒 × 12 家 = 每批白等 48 秒。
        deadline = time.time() + wait
        told = False
        first = True
        while True:
            for p in cands:
                if os.path.exists(p) and (first or self._file_settled(p)):
                    if told:
                        print("      内核就绪：%s" % p)
                    return p
            first = False
            if time.time() >= deadline:
                break
            if not told:
                print("    内核 %s 尚未就绪，等待紫鸟下载/更新（最多 %d 秒）…"
                      % (info.get("core_version") or "?", wait))
                print("      这在从未跑过紫鸟的机器上是正常的，只发生一次。")
                told = True
            time.sleep(5)

        raise ZiniaoError(
            "等了 %d 秒，店铺内核 %s 的 webdriver 仍未出现。找过：%s\n"
            "  如果这是新机器第一次跑：先手动打开紫鸟客户端、进一次这家店，"
            "让它把内核下完，再跑本脚本。\n"
            "  如果内核目录里确实没有 webdriver.exe（紫鸟安装不完整），"
            "用 python bootstrap.py --drivers 下载备用驱动。"
            % (wait, info.get("core_version"), "、".join(cands) or "（无候选路径）"))

    def _file_settled(self, path, settle=None):
        """文件大小在 settle 秒内没变化，且不为 0 —— 认为下载完成。"""
        settle = self.DRIVER_SETTLE if settle is None else settle
        try:
            a = os.path.getsize(path)
            if a == 0:
                return False
            time.sleep(settle)
            return os.path.getsize(path) == a
        except OSError:
            return False

    # startBrowser 回 -10000「发生未知错误」时，隔这么多秒再试，一共试这么多次
    # （合计约 5 分钟）。
    START_RETRY_WAITS = (15, 30, 45, 60, 60, 90)

    def _start_browser(self, request):
        """让紫鸟打开店铺。内核还在下载时它回 -10000，等一等再试。

        一家店第一次在这台机器上打开、或它的内核版本这台机器还没有时，紫鸟
        要先下内核（几百 MB）。这期间 startBrowser 不是等着，而是**一秒之内**
        回一个 statusCode -10000「发生未知错误」。实测两次都是这样：
        2026-10-12 BOCINA_TA02（内核 137）、2026-10-13 EWTTO_SM（内核 142），
        隔几十秒重跑就好了。12 家店用着至少三个内核版本，不重试的话新机器上
        第一轮会有好几家店白白失败。

        只对 -10000 重试：-10003（登录/权限）等别的错误再试也没用，照旧立刻抛。
        重试用尽还是 -10000，就按原来的方式抛出最后那次的错误。
        """
        waits = list(self.START_RETRY_WAITS)
        while True:
            resp = self.send(dict(request), timeout=300)
            if (resp is None or resp.get("statusCode") != -10000 or not waits):
                return self._require_ok(resp, "startBrowser")
            wait = waits.pop(0)
            print("  紫鸟回了「未知错误」(-10000) —— 多半是这家店的内核还在下载，"
                  "%d 秒后再试（还可再试 %d 次）" % (wait, len(waits) + 1))
            time.sleep(wait)

    @contextmanager
    def open_store(self, store, headless=False, close_when_done=True):
        oauth = store["browserOauth"] if isinstance(store, dict) else store
        name = store.get("browserName", oauth) if isinstance(store, dict) else oauth
        print("打开店铺：%s" % name)
        info = self._start_browser({
            "action": "startBrowser",
            "browserOauth": oauth,
            "isWaitPluginUpdate": 0,
            "isHeadless": 1 if headless else 0,
            "isWebDriverReadOnlyMode": 0,
            "cookieTypeLoad": 0,
            "cookieTypeSave": 0,
            "runMode": "1",
            "isLoadUserPlugin": False,
            "pluginIdType": 1,
            "privacyMode": 0,
        })

        # startBrowser 一成功，店铺浏览器就已经开着、占着它的代理了。下面
        # 找驱动、挂 webdriver 任何一步失败，都还没走到后面的 try/finally ——
        # 以前就这么把浏览器留在那里：内核没下完时每家店漏一个，批次后面的
        # 店铺接着在一堆没人管的窗口旁边开。所以这里自己关，再把异常放出去。
        step = "resolve_driver"
        try:
            drv_path = self._resolve_driver(info)
            opts = webdriver.ChromeOptions()
            opts.add_argument("--log-level=3")
            opts.add_experimental_option(
                "debuggerAddress", "127.0.0.1:%s" % info.get("debuggingPort"))
            step = "attach_webdriver"
            driver = webdriver.Chrome(service=Service(drv_path), options=opts)
            driver.implicitly_wait(20)
        except BaseException as e:
            try:
                if not getattr(e, "ziniao_step", None):
                    e.ziniao_step = step
            except Exception:
                pass
            print("  [失败] 店铺已启动但接不上（%s），关闭它" % step)
            self.close_store(info.get("browserOauth") or info.get("browserId")
                             or oauth)
            raise

        session = StoreSession(self, info, driver)
        self._opened.append(session)
        try:
            yield session
        finally:
            # never quit() - that would kill the browser we attached to.
            if close_when_done:
                self.close_store(session)
            self._stop_driver_service(driver)

    @staticmethod
    def _stop_driver_service(driver):
        """结束这家店的 chromedriver 进程。永不抛异常。

        因为不能 quit()，chromedriver 以前只靠 selenium 的 Service.__del__ 在
        解释器退出时收尸，而那一步在 Windows 上会以 WinError 6（句柄无效）失
        败，chromedriver 就留了下来。它继承了本进程的 stdout：上层（run_monthly、
        界面）读我们的输出要读到管道关闭为止，留一个 chromedriver 管道就永远
        不关，整条链卡死在"下载已经结束"之后。

        直接杀进程而不是 service.stop()：后者会先给 chromedriver 发 /shutdown，
        对一个 debuggerAddress 接上去的浏览器会做什么没有验证过；杀进程只是
        断开，浏览器（--keep-open 时还要留给人看）不受影响。
        """
        try:
            service = driver.service
            proc = service.process
            if proc is not None and proc.poll() is None:
                proc.kill()
                proc.wait(timeout=5)
            # 已经收拾干净，别让 Service.__del__ 在退出时再对着旧句柄报一屏错。
            service.process = None
        except Exception:
            pass

    def close_store(self, session):
        oauth = session.oauth if isinstance(session, StoreSession) else session
        print("关闭店铺 %s" % oauth)
        self.send({"action": "stopBrowser", "browserOauth": oauth, "duplicate": 0})
        if isinstance(session, StoreSession) and session in self._opened:
            self._opened.remove(session)

    def exit_client(self):
        self.send({"action": "exit"}, timeout=30)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False
