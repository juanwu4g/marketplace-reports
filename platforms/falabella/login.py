# -*- coding: utf-8 -*-
"""Get a logged-out Falabella store back in, with what Ziniao already filled in.

Since 2026-10 Falabella signs sellers in through its own single sign-on
(access-key-corp.falabella.tech, Keycloak):

  1. sellercenter.falabella.com/user/auth/login — email pre-filled, "Continuar"
  2. the sign-in page — email and password pre-filled, "Iniciar sesión"
  3. back to Seller Center

Step 2 has a bug that makes it reload forever on a slow connection: its inline
script adds ?fpkt=true to the address and calls location.reload() 500 ms later.
Through the store's proxy a page takes 0.7–0.9 s, so the reload always cancels
the new address and the script starts again (measured 2026-10-22 on
AIS_CL-Fala). While it loops the page never finishes loading and every Selenium
command waits. Opening the sign-in address with fpkt=true already in it stops
the loop; that navigation goes straight over the DevTools socket, because
chromedriver itself is stuck waiting for the looping page.

The owner approved pressing these two buttons (2026-10-22). Nothing is typed and
no password is read. One attempt per store; anything else on the way — an empty
email, an error, a password-reset page, a code or a captcha — stops here and
the store is reported as "needs login" (PAGE_BLOCKED), as before.
"""
import json
import time
import urllib.request

from platforms.falabella import page_kit as pk

LOGIN_PATH = "/user/auth/login"
SSO_HOST = "access-key-corp.falabella.tech"
SSO_AUTH = "/protocol/openid-connect/auth"
FIX = "fpkt=true"


def on_login_page(url):
    url = url or ""
    return LOGIN_PATH in url or SSO_HOST in url


def with_fix(url):
    """The sign-in address with fpkt=true, the parameter its own script adds."""
    if "fpkt=" in url:
        return url
    head, _, frag = url.partition("#")
    head += ("&" if "?" in head else "?") + FIX
    return head + ("#" + frag if frag else "")


# ------------------------------------------------------------------ DevTools socket

def _debugger(d):
    return (d.capabilities.get("goog:chromeOptions") or {}).get("debuggerAddress")


def _page_targets(addr):
    with urllib.request.urlopen("http://%s/json/list" % addr, timeout=5) as r:
        return [t for t in json.load(r) if t.get("type") == "page"]


def _sso_target(addr):
    for t in _page_targets(addr):
        if SSO_HOST in t.get("url", "") and SSO_AUTH in t.get("url", ""):
            return t
    return None


def _navigate(target, url):
    import websocket
    ws = websocket.create_connection(target["webSocketDebuggerUrl"], timeout=10,
                                     suppress_origin=True)
    try:
        ws.send(json.dumps({"id": 1, "method": "Page.navigate", "params": {"url": url}}))
        end = time.time() + 10
        while time.time() < end:
            if json.loads(ws.recv()).get("id") == 1:
                return
    finally:
        ws.close()


# ------------------------------------------------------------------ the two clicks

def sign_in(ctx, timeout=60):
    """Try once. Returns True when Seller Center is open again; otherwise
    False with the reason in ctx.flags["auto_login"] (the caller's page check
    then fails and the store is diagnosed from the page it is on)."""
    d, log = ctx.d, ctx.log
    ctx.flags["auto_login"] = "tried"
    log("  店铺没登录（Falabella 登录页）：用紫鸟已填好的账号自动登录一次")

    def stop(why):
        ctx.flags["auto_login"] = "failed: " + why
        log("  [自动登录没成功] %s —— 需要人工登录" % why)
        return False

    addr = _debugger(d)
    if not addr:
        return stop("拿不到浏览器调试端口")
    try:
        old_timeout = d.timeouts.page_load
    except Exception:
        old_timeout = 300
    d.set_page_load_timeout(20)       # a looping page must not hold a command for minutes
    try:
        if LOGIN_PATH in current_url(d):
            email = pk.find(d, "input[name=email]")
            if not email or not (email[0].get_attribute("value") or "").strip():
                return stop("登录页上没有预先填好的邮箱")
            go = pk.find(d, "input[type=submit], button[type=submit]")
            if not go:
                return stop("登录页上找不到「Continuar」")
            try:
                pk.real_click(d, go[0])
            except Exception:
                pass                      # the click navigates; chromedriver may time out
        target = None
        end = time.time() + timeout
        while time.time() < end and not target:
            target = _sso_target(addr)
            if not target:
                time.sleep(1)
        if not target:
            return stop("点了 Continuar 没有出现 Falabella 的统一登录页")
        _navigate(target, with_fix(target["url"]))
        time.sleep(3)
        if not pk.wait_for(lambda: pk.find(d, "input[type=password]"), timeout=30, poll=1):
            return stop("统一登录页没有密码框（可能要验证码或改密码）")
        user = pk.find(d, "input[name=username]")
        if user and not (user[0].get_attribute("value") or "").strip():
            return stop("统一登录页上邮箱是空的")
        btn = [b for b in pk.find(d, "button, input[type=submit]")
               if "iniciar" in ((b.text or b.get_attribute("value") or "").lower())]
        if not btn:
            return stop("找不到「Iniciar sesión」")
        try:
            pk.real_click(d, btn[0])     # a real click: the browser only releases an autofilled password to a real one
        except Exception:
            pass
        ok = pk.wait_for(lambda: _back_in(d), timeout=timeout, poll=2)
        if not ok:
            return stop("点了 Iniciar sesión 之后没有回到 Seller Center（停在 %s）"
                        % (current_url(d)[:80]))
        ctx.flags["auto_login"] = "ok"
        log("  已自动登录，回到 Seller Center")
        return True
    except Exception as e:
        return stop("%s: %s" % (e.__class__.__name__, str(e)[:120]))
    finally:
        try:
            d.set_page_load_timeout(old_timeout)
        except Exception:
            pass


def current_url(d):
    try:
        return d.current_url or ""
    except Exception:
        return ""


def _back_in(d):
    u = current_url(d)
    return "sellercenter.falabella.com" in u and not on_login_page(u)
