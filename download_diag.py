# -*- coding: utf-8 -*-
"""
没下到的报表，到底是为什么没下到。

以前一轮跑完只留下这样的话：

    mercadopago/account_statement 尚未就绪      （连着十一轮）
    已取回 0 张，仍欠 2 张
    [警告] 勾选了 8 张报表，实际得到 4 个文件

没有原因。可是绝大多数失败路径上，代码在放弃的那一刻是**知道**原因的 ——
行的状态、按钮是不是禁用、菜单里有什么、页面是不是被登录页顶掉了 —— 只是
打一行警告就把它扔了。这个模块做三件事：

  capture_evidence   放弃某张报表的当口，趁页面还在，把现场记下来
  classify           纯函数，不碰浏览器：证据 -> 一个封闭集合里的原因码
  expected_forms     这一趟**应该**产出哪些报表 —— 缺什么是算出来的，
                     而不是只对碰巧打出来的警告作反应

刻意不在模块顶层 import selenium / tkinter：classify 和 expected_forms 要能在
没有浏览器的机器上单测，报表端以后也可能只想读原因码表。
"""
import datetime
import io
import json
import os
import re
import time

import site_time

DIAG_DIR = "_diag"

# ------------------------------------------------------------------ 原因码
#
# 封闭集合。新增一种原因要在这里加一行，而不是在某个调用点临时拼一个字符串 ——
# 批次汇总按原因码计数，报表端按原因码决定要不要打扰人。
#
#   code: (一句话解释, 谁的问题, 过会儿重试有没有用, 操作员下一步)
CAUSES = {
    "STORE_OPEN_FAILED": (
        "店铺浏览器没打开，或中途断开", "紫鸟", True,
        "确认紫鸟客户端已登录、这家店能手动打开，再重跑这家店"),
    "LANGUAGE_WRONG": (
        "页面语言不是西班牙语，脚本按西语文字找不到控件", "账号", False,
        "在这家店的浏览器里把账号语言改回 Español，再重跑"),
    "NO_PERMISSION": (
        "这个登录账号没有查看报表的权限", "账号", False,
        "请账号管理员在 Colaboradores 里授予报表权限"),
    "PAGE_BLOCKED": (
        "打开的不是预期页面（登录、验证或错误页）", "账号", True,
        "手动进店完成登录/验证，确认页面能正常打开后重跑"),
    "UI_NOT_FOUND": (
        "页面是对的，但要点的控件找不到或点不了", "脚本", False,
        "把证据截图发给维护人：多半是平台改版，要改选择器"),
    "REQUEST_NOT_REGISTERED": (
        "点了生成，但平台的列表里没有这张报表", "平台", True,
        "重跑这家店重新申请；反复出现就在页面上手动生成一次"),
    "STILL_GENERATING": (
        "平台还在生成，尚未就绪", "平台", True,
        "稍后运行 run_batch.py --collect-only 补取"),
    "PLATFORM_DELAYED": (
        "平台提示数据还在处理，要等几个小时", "平台", True,
        "几小时后运行 run_batch.py --collect-only 补取"),
    "NO_DATA": (
        "该区间没有数据，平台不会生成报表", "平台", False,
        "无需处理"),
    "NOT_AVAILABLE": (
        "平台不提供：账期还没开始 / 没有这一期 / 店铺没有 Full", "平台", False,
        "无需处理；若认为应该有，对照证据核实"),
    "DOWNLOAD_NOT_LANDED": (
        "报表已就绪也点了下载，但文件没有落盘", "紫鸟", True,
        "重跑补取；反复出现就查磁盘空间、下载目录权限和浏览器的下载拦截"),
    "FILE_REJECTED": (
        "文件下到了但不能用（空文件，或已经不在目录里）", "平台", True,
        "重新下载这张报表"),
    "PLATFORM_ERROR": (
        "平台临时出错，返回的是 5xx 错误页", "平台", True,
        "过一会儿重跑这家店；持续出现就去平台状态页或手动打开页面确认"),
    "INTERNAL_ERROR": (
        "脚本自身出错", "脚本", True,
        "把异常信息发给维护人；可以先重跑一次"),
    "EXPIRED": (
        "超过保留期仍没取回，已不再追踪", "脚本", False,
        "还需要的话重新申请这张报表"),
    "UNKNOWN": (
        "现有证据分不清原因", "脚本", True,
        "打开证据文件和截图人工判断"),
}

# 这两种不需要任何人做任何事：缺了是对的。其余的都算"需要关注"。
NO_ATTENTION = ("NO_DATA", "NOT_AVAILABLE")

# 路线清单在这里抄了一份而不是 import mercadopago / meli_forms：那两个模块是
# 页面驱动层，这个模块要保持"不碰浏览器也能用"。tests 里有一条断言盯着两边
# 一致，谁加了报表类型忘了改这里会立刻红。
MP_REPORTS = ("settlement", "release", "collection", "withdraw",
              "after_collection", "account_statement")
STOCK_FORMS = ("stock/costos/1", "stock/costos/2", "stock/general",
               "stock/consolidado", "stock/devoluciones")

# 账单勾选框的 id -> 落盘文件名的开头。用来回答"勾了 4 张到了 3 个，缺的是哪张"。
BILLING_FILE_PREFIX = {
    "BILL_ML": "Reporte_Facturacion_MercadoLibre",
    "FULL": "Reporte_Cargos_Full",
    "PAYMENT": "Reporte_Pagos_Facturas",
    "NC_ML": "Reporte_Notas_Credito_MercadoLibre",
}

PERMISSION_WALL = "no tienes permisos para ver los reportes"

# 行不在列表里，要多久、探测几次，才敢说"平台没登记这次申请"。
ABSENT_AFTER_SECONDS = 15 * 60
ABSENT_MIN_PROBES = 2

# 折叠面板里这三种回答的意思都是"这个格式没人申请过"。
NOT_REGISTERED_ACTIONS = ("generar", "no-format-row", "no-panel")

# 登录 / 验证 / 浏览器错误页。按 URL 和标题判断，不看正文 —— 正文里出现
# "verifica" 的正常页面太多了。
BLOCKED_URL = re.compile(
    r"/login|/lgz/|/registration|/authentication|/verification|/captcha|"
    r"/security/|chrome-error://|/gz/account-verification|about:blank", re.I)
# 平台自己挂了，和"被挡在登录/验证页外面"不是一回事：前者谁也不用去处理，等
# 一会儿重跑就行。2026-10-13 实测 MercadoPago 的页面标题就是
# 「503 Service Temporarily Unavailable」，当时被归成了要人工进店验证。
SERVER_ERROR = re.compile(
    r"\b50[0234]\b|service (temporarily )?unavailable|bad gateway|"
    r"gateway time-?out|internal server error", re.I)
BLOCKED_TITLE = re.compile(
    r"ingresa a tu cuenta|iniciar sesi[oó]n|inicia sesi[oó]n|\blog ?in\b|"
    r"verifica|captcha|no es posible acceder|p[aá]gina no encontrada|"
    r"\b40[34]\b|\b50[023]\b|ERR_|no se puede acceder", re.I)

# 浏览器没了。和"脚本出错"不是一回事：重开店铺就好。
DEAD_SESSION = re.compile(
    r"invalid session id|no such window|chrome not reachable|disconnected|"
    r"target window already closed|session deleted|web view not found|"
    r"connection refused|max retries exceeded", re.I)

# 点了下载什么都没来，而页面上弹出的提示条写着"没有数据"。只认提示条里的
# 字，不认正文：正文里"no hay ..."到处都是。措辞是按平台已知的几句话写的
# （"no hay movimientos en este rango de fechas"），没有在库存报表页上实测过。
NO_DATA_ALERT = re.compile(
    r"no hay (movimientos|datos|informaci[oó]n|devoluciones|registros)|"
    r"no (tienes|tenemos|encontramos) (movimientos|datos|devoluciones|registros)|"
    r"sin (movimientos|datos|devoluciones)", re.I)

UI_EXCEPTION_CLASSES = (
    "NoSuchElementException", "TimeoutException",
    "ElementNotInteractableException", "ElementClickInterceptedException",
    "StaleElementReferenceException", "NoSuchShadowRootException")
UI_EXCEPTION_TEXT = re.compile(
    r"not found|never rendered|never enabled|找不到|未找到", re.I)

SIGNALS = (
    "store-open", "expired", "exception", "no-permission",
    "ui-missing", "ui-disabled", "request-none", "not-ready",
    "download-not-landed", "short-download", "file-rejected",
    "wrong-period-page", "no-month-on-page", "not-offered",
    "period-not-started", "not-attempted")


# ------------------------------------------------------------------ 小工具

def _iso(ts):
    """epoch 秒 -> 'YYYY-MM-DDTHH:MM:SS'；本来就是字符串或 None 的原样返回。"""
    if ts is None or isinstance(ts, str):
        return ts
    try:
        return datetime.datetime.fromtimestamp(float(ts)).isoformat(
            timespec="seconds")
    except Exception:
        return None


def _plain(v, depth=0):
    """压成能进 JSON 的东西。证据里什么都可能被塞进来（WebElement、异常对象），
    而证据文件写不出来不该反过来把下载搞挂。"""
    if v is None or isinstance(v, (bool, int, float)):
        return v
    if isinstance(v, str):
        return v if len(v) <= 500 else v[:500] + "…"
    if depth > 4:
        return repr(v)[:200]
    if isinstance(v, dict):
        return {str(k): _plain(x, depth + 1) for k, x in list(v.items())[:60]}
    if isinstance(v, (list, tuple, set, frozenset)):
        seq = sorted(v, key=repr) if isinstance(v, (set, frozenset)) else v
        return [_plain(x, depth + 1) for x in list(seq)[:60]]
    return repr(v)[:200]


def _safe_name(label):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(label or "form")).strip("_")[:60]


def exception_evidence(exc, step=None):
    """一个异常能留下的全部：类名、文字、哪一步，以及紫鸟的原始回应（若有）。"""
    ev = {"signal": "exception",
          "exception_class": exc.__class__.__name__,
          "exception": str(exc)[:500]}
    step = step or getattr(exc, "ziniao_step", None)
    if step:
        ev["step"] = step
    resp = getattr(exc, "ziniao_response", None)
    if resp:
        ev["ziniao_response"] = str(resp)[:500]
    return ev


def folder_delta(before, after):
    """点了下载之后目录里多了什么。

    `.crdownload` 单列：它说明下载**开始过**但没写完 —— 和"点了没反应"是两种
    毛病，前者查磁盘和网络，后者查页面。
    """
    new = sorted(set(after) - set(before or ()))
    new = [n for n in new if n != DIAG_DIR]
    half = [n for n in new if n.endswith((".crdownload", ".tmp"))]
    return {"folder_delta": [n for n in new if n not in half],
            "crdownload": half}


def dedupe(seq):
    """去重，保持先后顺序。"""
    seen, out = set(), []
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def match_billing_files(checked_ids, files):
    """勾选的报表 id 对上落盘的文件，返回 (对上的 {id: 文件}, 缺的 id, 认不出的 id)。

    认不出的 id（平台新加的报表类型）单独返回：不知道它该叫什么文件名，就不能
    说它缺了，也不能说它到了。
    """
    got, missing, unknown = {}, [], []
    left = list(files)
    for cid in checked_ids:
        prefix = BILLING_FILE_PREFIX.get(cid)
        if not prefix:
            unknown.append(cid)
            continue
        hit = next((f for f in left
                    if os.path.basename(f).lower().startswith(prefix.lower())),
                   None)
        if hit is None:
            missing.append(cid)
        else:
            got[cid] = hit
            left.remove(hit)
    return got, missing, unknown


# ------------------------------------------------------------------ 取证

# 只读，不点任何东西。可见性按包围盒判断，和下载层其余地方一个口径。
JS_PAGE_FACTS = """
const wall = arguments[0];
const txt = (s, n) => (s || '').replace(/\\s+/g, ' ').trim().slice(0, n);
const vis = (e) => { const r = e.getBoundingClientRect();
                     return r.width > 3 && r.height > 3; };
const full = document.body ? (document.body.innerText || '') : '';
const alerts = [];
for (const e of document.querySelectorAll(
    '[role=alert], [role=alertdialog], .andes-message, .andes-snackbar, ' +
    '.andes-ui-message, .andes-ui-snackbar, .ui-message')) {
  if (!vis(e)) continue;
  const t = txt(e.innerText, 240);
  if (t && alerts.indexOf(t) < 0) alerts.push(t);
  if (alerts.length >= 6) break;
}
return {lang: document.documentElement.getAttribute('lang') || '',
        body: txt(full, 300),
        permission_wall: full.toLowerCase().indexOf(wall) >= 0,
        alerts: alerts};
"""


def capture_evidence(driver, label, out_dir, extra=None, shot=True):
    """放弃一张报表的当口，把现场记下来。返回一个普通 dict。

    三条硬要求：

      * 趁页面还在时调用 —— 之后的 driver.get 会把现场冲掉
      * 便宜：一次属性读取、一次 JS、一张截图，一秒上下。不做任何等待
      * **绝不抛异常**。取证是为了解释失败，它自己失败不能变成新的失败 ——
        每一步各自兜住，拿到多少算多少，拿不到的记在 capture_errors 里

    `driver` 可以是 None（店铺根本没打开）：那就只留 `extra` 里带来的东西。
    `shot=False` 不截图也不落盘，给轮询用 —— 每一轮都截图毫无意义，只有最后
    一次探测的画面才是"放弃时的现场"。
    """
    ev = {"label": label, "captured_at": _iso(time.time()),
          "url": None, "title": None, "lang": None, "body": None,
          "alerts": [], "screenshot": None, "file": None}
    errors = []
    try:
        for k, v in (extra or {}).items():
            ev[str(k)] = _plain(v)
    except Exception as e:
        errors.append("extra: %s" % str(e)[:80])

    if driver is None:
        ev["no_driver"] = True
    else:
        try:
            ev["url"] = _plain(driver.current_url)
        except Exception as e:
            errors.append("url: %s" % str(e)[:80])
        try:
            ev["title"] = _plain(driver.title)
        except Exception as e:
            errors.append("title: %s" % str(e)[:80])
        try:
            facts = driver.execute_script(JS_PAGE_FACTS, PERMISSION_WALL) or {}
            ev["lang"] = _plain(facts.get("lang"))
            ev["body"] = _plain(facts.get("body"))
            ev["alerts"] = _plain(facts.get("alerts") or [])
            if facts.get("permission_wall"):
                ev["permission_wall"] = True
        except Exception as e:
            errors.append("page: %s" % str(e)[:80])

    if shot:
        try:
            folder = os.path.join(out_dir, DIAG_DIR)
            if not os.path.isdir(folder):
                os.makedirs(folder)
            stem = "%s_%s" % (datetime.datetime.now().strftime("%H%M%S"),
                              _safe_name(label))
            if driver is not None:
                try:
                    png = os.path.join(folder, stem + ".png")
                    if driver.save_screenshot(png):
                        ev["screenshot"] = "%s/%s.png" % (DIAG_DIR, stem)
                except Exception as e:
                    errors.append("screenshot: %s" % str(e)[:80])
            ev["file"] = "%s/%s.json" % (DIAG_DIR, stem)
            if errors:
                ev["capture_errors"] = errors
            with io.open(os.path.join(folder, stem + ".json"), "w",
                         encoding="utf-8") as fh:
                fh.write(json.dumps(ev, ensure_ascii=False, indent=2))
        except Exception as e:
            ev["file"] = None
            errors.append("save: %s" % str(e)[:80])
    if errors:
        ev["capture_errors"] = errors
    return ev


# ------------------------------------------------------------------ 归因

def _server_error(ev):
    title = ev.get("title") or ""
    if title and SERVER_ERROR.search(title):
        return "标题是「%s」" % title[:60]
    body = (ev.get("body_text") or ev.get("body") or "")[:120]
    if body and SERVER_ERROR.search(body) and len(body.strip()) < 120:
        return "页面只有一句「%s」" % body.strip()[:60]
    return None


def _blocked(ev):
    url, title = ev.get("url") or "", ev.get("title") or ""
    if url and BLOCKED_URL.search(url):
        return "地址是 %s" % url[:80]
    if title and BLOCKED_TITLE.search(title):
        return "标题是「%s」" % title[:60]
    return None


def _no_permission(ev):
    if ev.get("permission_wall") or ev.get("signal") == "no-permission":
        return True
    if ev.get("exception_class") == "NoReportAccess":
        return True
    hay = " ".join([ev.get("body") or "", ev.get("exception") or ""]
                   + list(ev.get("alerts") or [])).lower()
    return PERMISSION_WALL in hay


def _absent(ev, ctx, what):
    """列表里没有这一行。够久、够多次才敢说"没登记"，否则老实说分不清。"""
    probes = ctx.get("absent_probes") or 0
    asked = ctx.get("requested_at")
    last = ctx.get("last_seen")
    if last is None:
        last = ctx.get("now")
    waited = (last - asked) if (asked is not None and last is not None) else None
    if (probes >= ABSENT_MIN_PROBES and waited is not None
            and waited > ABSENT_AFTER_SECONDS):
        return "REQUEST_NOT_REGISTERED", (
            "申请后 %d 分钟、探测 %d 次，%s始终没有出现"
            % (waited // 60, probes, what))
    return "UNKNOWN", (
        "%s没有出现，但%s，分不清是平台还没列出来还是根本没登记这次申请"
        % (what,
           "只探测到 %d 次、距申请 %s" % (
               probes, "%d 分钟" % (waited // 60) if waited is not None
               else "多久未知")))


def _decide(ev, ctx):
    """证据 -> (原因码, 补充说明)。规则按从上到下的顺序生效，先中先得。

    顺序本身就是规则的一部分：越是"一锤定音"的事实排得越前。页面被登录页
    顶掉时，后面关于按钮、行状态的证据全是对着错误页面采的，不能信。
    """
    signal = ev.get("signal")
    exc_text = ev.get("exception") or ""
    exc_class = ev.get("exception_class") or ""
    row = ev.get("row_state")

    # 1. 过期、店铺没打开：和页面无关
    if signal == "expired":
        return "EXPIRED", "申请于 %s" % (_iso(ctx.get("requested_at")) or "未知时间")
    if signal == "store-open":
        return "STORE_OPEN_FAILED", "%s%s%s" % (
            "卡在 %s；" % ev["step"] if ev.get("step") else "",
            ("%s: %s" % (exc_class, exc_text))[:160] if exc_text else "没有异常信息",
            "；紫鸟回应 %s" % ev["ziniao_response"][:120]
            if ev.get("ziniao_response") else "")
    if DEAD_SESSION.search(exc_text):
        return "STORE_OPEN_FAILED", "浏览器中途断开：%s" % exc_text[:120]

    # 2. 页面不是我们以为的那个：登录/验证/错误页，或没有权限
    why = _server_error(ev)
    if why:
        return "PLATFORM_ERROR", why
    why = _blocked(ev)
    if why:
        return "PAGE_BLOCKED", why
    if _no_permission(ev):
        return "NO_PERMISSION", "页面写着 No tienes permisos para ver los reportes"

    # 3. 平台明说的终态。读的是行的 class，不依赖语言
    if row == "empty":
        return "NO_DATA", ev.get("status_text") or ""
    if row == "delayed":
        return "PLATFORM_DELAYED", ev.get("status_text") or ""

    # 4. 文件层面
    if signal == "file-rejected":
        return "FILE_REJECTED", "、".join(
            "%s" % f for f in (ev.get("files") or [])[:4])

    # 5. 平台不提供
    if signal == "period-not-started":
        return "NOT_AVAILABLE", "账期还没开始"
    if signal == "not-offered":
        return "NOT_AVAILABLE", ev.get("step") or "平台只提供更少的期数"
    if ctx.get("has_full") is False and str(ctx.get("form", "")).startswith("stock/"):
        return "NOT_AVAILABLE", "这家店没有 Full 仓"
    if signal == "wrong-period-page":
        detail = "页面显示的是 %s，不是 %s" % (ev.get("shown"), ev.get("want"))
        if ctx.get("required"):
            # 目标月本身拿不到，不能当成"平台不提供"轻轻放过 —— 那样整家店
            # 会显示 ok，而它要的那个月一张账单都没有。
            return "PAGE_BLOCKED", detail + "（目标账期被重定向走了）"
        return "NOT_AVAILABLE", detail + "（后续账期不可访问）"

    # 6. 就绪了、点了、文件没到。和语言无关：能判出"就绪"说明页面读对了
    if signal == "download-not-landed":
        said = next((a for a in ev.get("alerts") or []
                     if NO_DATA_ALERT.search(a or "")), None)
        if said:
            return "NO_DATA", "页面提示：%s" % said[:120]
    if (signal == "short-download" and ev.get("unrecognised_ids")
            and not ev.get("missing_reports")):
        # 认识的报表都到了，少的是脚本不认识的那种。这不是"文件没落盘"：
        # 2026-10-13 EWTTO_SM 的取证截图里，平台为 BILL_MP_V3（Mercado Pago
        # 账单）弹了一个"选 XLSX 还是 CSV"的对话框等人回答，脚本没有处理它，
        # 那张报表于是从来没下过。归到"控件没处理"，由维护人补上这一步。
        return "UI_NOT_FOUND", (
            "勾选了脚本不认识的报表类型 %s，没有对应的文件（平台可能为它另弹了"
            "对话框，脚本没有处理 —— 看证据截图）"
            % "、".join(ev["unrecognised_ids"]))
    if signal in ("download-not-landed", "short-download"):
        half = ev.get("crdownload") or []
        detail = ("下载开始过但没写完：%s" % "、".join(half[:3]) if half
                  else "点击后没有任何下载开始")
        if signal == "short-download":
            detail = "勾选 %s，缺 %s" % (
                len(ev.get("checked") or []) or "?",
                "、".join(ev.get("missing_reports") or []) or "（对不上是哪张）")
        return "DOWNLOAD_NOT_LANDED", detail

    # 7. 语言。排在控件类规则之前：语言错了，"找不到按钮"只是症状
    lang = (ev.get("lang") or "").strip().lower()
    if lang and not lang.startswith("es"):
        return "LANGUAGE_WRONG", "<html lang=%s>" % lang
    if "界面语言" in exc_text:
        return "LANGUAGE_WRONG", exc_text[:120]

    # 8. 异常：是页面控件的事，还是脚本的事
    if signal == "exception" or exc_class:
        text = ("%s: %s" % (exc_class, exc_text))[:200]
        if exc_class in UI_EXCEPTION_CLASSES or (
                exc_class == "RuntimeError" and UI_EXCEPTION_TEXT.search(exc_text)):
            return "UI_NOT_FOUND", text
        return "INTERNAL_ERROR", text

    # 9. 控件
    if signal in ("ui-missing", "ui-disabled"):
        return "UI_NOT_FOUND", "%s%s" % (
            ev.get("step") or "",
            "（控件在但被禁用）" if signal == "ui-disabled" else "")
    if signal == "no-month-on-page":
        return "PAGE_BLOCKED", "账单明细页上没有可识别的月份"
    if row == "ready":
        # 判了就绪却没有可点的入口
        return "UI_NOT_FOUND", "行已就绪，但没有可点击的下载入口"

    # 10. 还没到：区分"在生成"和"压根没登记"
    if row == "accordion":
        action = ev.get("action")
        if action in NOT_REGISTERED_ACTIONS:
            return "REQUEST_NOT_REGISTERED", "折叠面板里该格式是 %s" % action
        if action in (None, "", "absent"):
            return _absent(ev, ctx, "该月的折叠面板")
        if action == "no-link" or re.search(r"prepar|proces|generando", action, re.I):
            return "STILL_GENERATING", "折叠面板里该格式是 %s" % action
        return "UNKNOWN", "折叠面板里该格式显示 %r，不是已知的任何一种状态" % action
    if row == "absent":
        return _absent(ev, ctx, "这一行")
    if row in ("generating", "waiting"):
        return "STILL_GENERATING", ev.get("status_text") or ""
    if row == "ventas-widget":
        n, before = ev.get("widget_rows") or 0, ev.get("rows_before") or 0
        sub = ev.get("top_sub") or ""
        if not n:
            return _absent(ev, ctx, "通知小组件里的记录")
        if "xito" in sub and ev.get("top_has_dl") and n <= before:
            return "UNKNOWN", (
                "小组件最上面一条是成功记录，但行数没有比申请前多（%d → %d），"
                "分不清它是这次的还是以前的" % (before, n))
        if re.search(r"error|no pudimos|fall", sub, re.I):
            return "UNKNOWN", "小组件最上面一条写着「%s」" % sub[:80]
        return "STILL_GENERATING", sub[:80]

    # 11. 兜底。必须带着证据走，不许空手说"不知道"
    if signal == "request-none":
        return "UNKNOWN", "申请函数没有返回任何东西，也没有留下更具体的线索"
    if signal == "not-attempted":
        return "UNKNOWN", "前一期失败后脚本没有再试这一期"
    return "UNKNOWN", "没有能定性的证据（signal=%s, row_state=%s）" % (signal, row)


def classify(evidence, context=None):
    """证据 + 上下文 -> 一条诊断记录。纯函数，不碰浏览器，不读时钟。

    context: form, period, required, requested_at, probes, absent_probes,
             first_seen, last_seen, now, has_full
    """
    ev = dict(evidence or {})
    ctx = dict(context or {})
    try:
        code, detail = _decide(ev, ctx)
    except Exception as e:                 # 归因出错也要有个交代
        code, detail = "UNKNOWN", "归因规则自身出错：%s" % str(e)[:120]
    if code == "UNKNOWN" and not ev:
        detail = "没有采到任何证据"
    explain, owner, retry, action = CAUSES[code]
    return {"form": ctx.get("form"), "period": ctx.get("period") or "",
            "cause": code,
            "explain": explain + ("：%s" % detail if detail else ""),
            "owner": owner, "retry": retry, "action": action,
            "evidence": ev,
            "first_seen": _iso(ctx.get("first_seen")),
            "last_seen": _iso(ctx.get("last_seen")),
            "probes": ctx.get("probes") or 0}


# 要人动手才能好的原因：(标题, 怎么处理, 处理完重跑有没有用)。
# 这几种脚本自己永远解决不了 —— 不让人一眼看到，就会每个月原样失败一次。
HUMAN = {
    "PAGE_BLOCKED": ("要登录 / 验证",
                     "点「打开所选店铺」，在弹出的店铺窗口里按页面提示完成登录或验证",
                     True),
    "LANGUAGE_WRONG": ("账号语言不对",
                       "点「打开所选店铺」，把账号语言改回 Español", True),
    "STORE_OPEN_FAILED": ("店铺打不开",
                          "确认紫鸟客户端已登录、这家店在紫鸟里能手动打开", True),
    "NO_PERMISSION": ("缺 MercadoPago 报表权限",
                      "请账号管理员在 Colaboradores 里开通报表权限；开通之前重跑也没用",
                      False),
}


def human_actions(rows):
    """批次汇总里需要人处理的店。纯函数。

    rows 是 logs/batch_<stamp>.json 的 "stores"。同一家店同一种原因并成一条：
    [{"store", "cause", "title", "how", "rerun", "forms", "evidence"}]。
    只看这一趟自己的缺失（"missing"），更早几趟留下的（"carried"）不算。
    """
    out = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        by = {}
        for r in row.get("missing") or []:
            if isinstance(r, dict) and r.get("cause") in HUMAN:
                by.setdefault(r["cause"], []).append(r)
        for cause, recs in by.items():
            title, how, rerun = HUMAN[cause]
            out.append({"store": row.get("store"), "cause": cause,
                        "title": title, "how": how, "rerun": rerun,
                        "forms": [r.get("form") for r in recs],
                        "evidence": evidence_ref(recs[0]) or ""})
    return out


# 下载了、但财务报表（report_build.py）一张都不读的报表。缺了照样记原因、写进
# REPORT.md，但不算"报表会缺数据" —— 不让它们触发告警。2026-10-14 那轮的 6 次
# "部分报表没下成"里有 4 次只缺这几种（Retiros 的 Generar 点不了、Poscobro /
# 账户对账单碰上 503），报表本身一个数都不受影响。
# 报表读的是：Ventas、账单 4 张、Full 费用、仓储费、库存总表、退货质检、
# settlement 和 reserve-release。改 report_build 读哪些文件时同步改这里。
NOT_IN_REPORT = {
    "mercadopago/collection", "mercadopago/withdraw",
    "mercadopago/after_collection", "mercadopago/account_statement",
    "stock/consolidado",
}


def affects_report(record):
    """缺了这张，财务报表会不会少数据。"""
    form = (record or {}).get("form") or ""
    return form not in NOT_IN_REPORT


def needs_attention(record):
    return record.get("cause") not in NO_ATTENTION


def count_by_cause(records):
    out = {}
    for r in records:
        out[r.get("cause")] = out.get(r.get("cause"), 0) + 1
    return out


# ------------------------------------------------------------------ 应有的报表

def _month_of(args):
    v = getattr(args, "month", None)
    if not v:
        return None
    try:
        return int(v[:4]), int(v[5:7])
    except Exception:
        return None


def expected_forms(args, today=None):
    """这一趟按参数**应该**产出的报表 id 列表。

    id 的写法：
        ventas
        billing/2026-07          --month 路线，按会计月
        billing/1, billing/2     常规路线，第 N 新的账期（事先不知道月份名）
        stock/costos/1 ...       仓储费按账期，最新的在前
        mercadopago/<类型>

    还没开始的账期不算在内 —— download_billing_for_months 本来就不去请求它，
    把它列成"缺失"只会让每家店每一趟都多出一行没人需要看的东西。
    """
    only = getattr(args, "only", "all") or "all"
    today = today or site_time.site_today()
    out = []
    if only in ("ventas", "all"):
        out.append("ventas")
    if only in ("mercadopago", "all"):
        out.extend("mercadopago/%s" % k for k in MP_REPORTS)
    if only in ("billing", "all"):
        want = _month_of(args)
        if want:
            y, mo = want
            for _ in range(max(1, (getattr(args, "billing_after", 2) or 0) + 1)):
                if datetime.date(y, mo, 1) > today:
                    break
                out.append("billing/%04d-%02d" % (y, mo))
                mo += 1
                if mo > 12:
                    y, mo = y + 1, 1
        else:
            out.extend("billing/%d" % (i + 1)
                       for i in range(max(0, getattr(args, "months", 2) or 0)))
    if only in ("stock", "all"):
        out.extend(STOCK_FORMS)
    return out


def route_of(form):
    return str(form or "").split("/")[0]


# ------------------------------------------------------------------ 收集器

class Collector(object):
    """一家店一个。下载层把"放弃了什么、当时看到什么"交到这里。

    为什么是显式传进去的对象，而不是 meli_forms.SALES_PERIOD_ACTUAL 那样的
    模块级变量：批量跑时模块级状态会带着上一家店的值进下一家 —— 那两个变量
    已经各出过一次串店的事。诊断要是串了店，比没有诊断更糟。
    """

    def __init__(self, out_dir, store=None):
        self.out_dir = out_dir
        self.store = store
        self.items = {}        # form -> {"evidence", "period", "required"}
        self.delivered = {}    # form -> {"period", "files"}
        self.flags = {}
        self.cautions = []

    def give_up(self, driver, form, signal, period=None, step=None,
                required=None, shot=True, **extra):
        """记下一次放弃。同一张报表只留**第一次**的现场。

        后来的放弃多半是第一次的后果（按钮禁用 -> 之后自然也没有文件），
        只在第一条证据的 later 里留个名字。
        """
        info = dict(extra)
        info["signal"] = signal
        if step:
            info["step"] = step
        if form in self.items:
            later = self.items[form]["evidence"].setdefault("later", [])
            if len(later) < 5:
                later.append({"signal": signal, "step": step})
            return self.items[form]["evidence"]
        ev = capture_evidence(driver, form, self.out_dir, info, shot=shot)
        self.items[form] = {"evidence": ev, "period": period,
                            "required": required}
        return ev

    def mark_delivered(self, form, files, period=None):
        self.delivered[form] = {"period": period, "files": list(files or [])}

    def flag(self, name, value):
        # dict 是合并不是覆盖：range_clamped 可能来自销售窗口，也可能来自
        # 库存报表的日期区间，后到的不能把先到的冲掉
        if isinstance(value, dict) and isinstance(self.flags.get(name), dict):
            self.flags[name].update(_plain(value))
        else:
            self.flags[name] = _plain(value)

    def caution(self, form, kind, **info):
        c = {"form": form, "kind": kind}
        c.update(_plain(info))
        self.cautions.append(c)


# 下面三个是给下载层用的入口：diag 为 None 时什么都不做（download_history.py
# 等老调用方不传），并且**绝不抛异常** —— 诊断不能拖垮它要诊断的下载。

def give_up(diag, driver, form, signal, **kw):
    if diag is None:
        return None
    try:
        return diag.give_up(driver, form, signal, **kw)
    except Exception as e:
        print("  [警告] 诊断取证失败（%s）：%s" % (form, str(e)[:80]))
        return None


def delivered(diag, form, files, period=None):
    if diag is None or not files:
        return
    try:
        diag.mark_delivered(form, files, period)
    except Exception:
        pass


def flag(diag, name, value):
    if diag is None:
        return
    try:
        diag.flag(name, value)
    except Exception:
        pass


# ------------------------------------------------------------------ 待取条目

def note_probe(entry, info, now=None):
    """把一次探测的结果折进待取条目。只留最新一次的状态 + 计数 + 首末时间。

    条目会原样写进 logs/pending.json，所以这里只放纯数据。
    """
    now = time.time() if now is None else now
    info = dict(info or {})
    entry["probes"] = (entry.get("probes") or 0) + 1
    entry.setdefault("first_probe_at", now)
    entry["last_probe_at"] = now
    state = info.get("row_state") or "unknown"
    entry["last_state"] = state
    entry["last_status_text"] = info.get("status_text") or info.get("top_sub") or ""
    absent = state == "absent" or (
        state == "accordion" and info.get("action") in (None, "", "absent")) or (
        state == "ventas-widget" and not info.get("widget_rows"))
    # 连续不在才算数：中间出现过一次，就说明平台是登记了的
    entry["absent_probes"] = (entry.get("absent_probes") or 0) + 1 if absent else 0
    ev = dict(entry.get("evidence") or {})
    # 上一轮的专属字段不能留到这一轮：行出现之后还带着 action/formats 的旧值，
    # 归因会被过期的事实带偏
    for k in ("row_state", "status_text", "formats", "action", "head", "cells",
              "widget_rows", "rows_before", "top_sub", "top_has_dl", "signal",
              "exception", "exception_class", "folder_delta", "crdownload",
              "fetch", "step", "ziniao_response") + PAGE_KEYS:
        ev.pop(k, None)
    ev.update(_plain(info))
    entry["evidence"] = ev
    return entry


# capture_evidence 自己负责的键。重新取证时不能把上一次的带进去，否则这次
# 截图失败的话，条目上留着的就是一张指向旧画面的截图。
PAGE_KEYS = ("label", "captured_at", "url", "title", "lang", "body", "alerts",
             "screenshot", "file", "permission_wall", "capture_errors",
             "no_driver")


def snapshot_entry(driver, entry, out_dir, shot=True):
    """给待取条目拍下"放弃时的现场"。只在最后一次探测时调用。绝不抛异常。"""
    try:
        extra = {k: v for k, v in (entry.get("evidence") or {}).items()
                 if k not in PAGE_KEYS}
        entry["evidence"] = capture_evidence(
            driver, form_of_entry(entry), out_dir, extra, shot=shot)
    except Exception as e:
        print("  [警告] 诊断取证失败（%s）：%s" % (form_of_entry(entry), str(e)[:80]))
    return entry


def failure_signal(exc):
    """异常是"店铺没打开"还是别的。紫鸟报的错、以及 open_store 里标了
    ziniao_step 的，都算前者。"""
    if exc.__class__.__name__ == "ZiniaoError" or getattr(exc, "ziniao_step", None):
        return "store-open"
    return "exception"


def note_failure(entry, exc, now=None):
    """补收时店铺没打开（或中途出错）：条目照旧欠着，把异常记成它的证据。

    不算一次探测 —— 根本没看到页面。之前的行状态作废：那是上一趟看到的，
    留着会让归因把这次的失败说成"还在生成"。
    """
    signal = failure_signal(exc)
    ev = dict(exception_evidence(exc), signal=signal, no_driver=True,
              row_state="probe-error",
              captured_at=_iso(time.time() if now is None else now))
    entry["evidence"] = ev
    entry["last_state"] = ("store-open-failed" if signal == "store-open"
                           else "probe-error")
    entry["last_status_text"] = ("%s: %s" % (exc.__class__.__name__, exc))[:160]
    return entry


def store_failure_records(exc, expected):
    """整家店没跑成：应有的每张报表一条记录，原因相同。"""
    ev = dict(exception_evidence(exc), signal=failure_signal(exc), no_driver=True)
    return [classify(ev, {"form": f}) for f in expected]


def form_of_entry(entry):
    if entry.get("kind") == "ventas":
        return "ventas"
    return "mercadopago/%s" % entry.get("report", "?")


def record_from_entry(entry, now=None):
    """待取条目 -> 诊断记录。用的是条目上最后一次探测留下的证据。"""
    ev = dict(entry.get("evidence") or {})
    if "row_state" not in ev and entry.get("last_state"):
        ev["row_state"] = entry["last_state"]
    if "status_text" not in ev and entry.get("last_status_text"):
        ev["status_text"] = entry["last_status_text"]
    if entry.get("adopted_existing_row"):
        ev["adopted_existing_row"] = entry["adopted_existing_row"]
    ctx = {"form": form_of_entry(entry),
           "period": entry.get("range") or "",
           "requested_at": entry.get("requested_at"),
           "probes": entry.get("probes") or 0,
           "absent_probes": entry.get("absent_probes") or 0,
           "first_seen": entry.get("first_probe_at"),
           "last_seen": entry.get("last_probe_at"),
           "now": time.time() if now is None else now}
    rec = classify(ev, ctx)
    rec["requested_at"] = _iso(entry.get("requested_at"))
    return rec


def expired_record(entry):
    ev = dict(entry.get("evidence") or {})
    ev["signal"] = "expired"
    rec = record_from_entry(dict(entry, evidence=ev))
    rec["store"] = entry.get("store")
    rec["out_dir"] = entry.get("out_dir")
    return rec


# ------------------------------------------------------------------ 诊断这一步

def _rejected(files, out_dir):
    """记在账上的文件，是不是真的在、真的有内容。"""
    bad = []
    for f in files or []:
        p = os.path.join(out_dir, os.path.basename(f))
        try:
            if os.path.getsize(p) == 0:
                bad.append("%s（0 字节）" % os.path.basename(f))
        except OSError:
            bad.append("%s（不在目录里）" % os.path.basename(f))
    return bad


def diagnose_store(expected, collector=None, entries=(), no_data=(),
                   collected=None, out_dir=None, now=None, has_full=None):
    """一家店跑完：应有的报表里哪些没到，各是为什么。返回记录列表。

    expected   expected_forms() 的结果
    collector  这家店的 Collector（同步路线和申请失败的证据在里面）
    entries    还欠着的异步报表条目（collect_pending 写过诊断字段）。这些记录
               带 owed=True：它们的结论会在补收之后被 merge_final 换掉
    no_data    平台明说该区间没有数据的条目，终态
    collected  {报表 id: [文件]}，异步路线已经取回的
    """
    now = time.time() if now is None else now
    items = collector.items if collector is not None else {}
    done = dict(collected or {})
    if collector is not None:
        for form, d in collector.delivered.items():
            done[form] = d.get("files") or []
    by_form, owed = {}, set()
    for e in no_data or ():
        by_form[form_of_entry(e)] = e
    for e in entries or ():
        by_form[form_of_entry(e)] = e
        owed.add(form_of_entry(e))

    out = []

    def add(rec):
        if not any(r["form"] == rec["form"] for r in out):
            out.append(rec)

    for form in expected:
        if form in done:
            bad = _rejected(done[form], out_dir) if out_dir else []
            if bad:
                add(classify({"signal": "file-rejected", "files": bad},
                             {"form": form, "now": now}))
            continue
        if form in by_form:
            rec = record_from_entry(by_form[form], now=now)
            if form in owed:
                rec["owed"] = True
            add(rec)
            continue
        item = items.get(form)
        inherited = False
        if item is None and route_of(form) in items:
            # 整条路线一起倒下的（异常、没有权限）：这张报表自己没有现场，
            # 用路线那一份。
            item, inherited = items[route_of(form)], True
        if item is None:
            add(classify({}, {"form": form, "now": now, "has_full": has_full}))
            continue
        ev = dict(item["evidence"])
        if inherited:
            ev["inherited_from"] = route_of(form)
        add(classify(ev, {"form": form, "period": item.get("period"),
                          "required": item.get("required"),
                          "now": now, "has_full": has_full}))

    # 不在应有清单里、但下载层明确放弃了的子项：账单某一期勾了 4 张只到 3 张，
    # 缺的那张是 billing/1/FULL —— 这一期算"到了"，那一张不算。
    for form, item in items.items():
        if form in expected or "/" not in form or form in done:
            continue
        add(classify(dict(item["evidence"]),
                     {"form": form, "period": item.get("period"),
                      "required": item.get("required"), "now": now}))

    # "没试"的跟着害它没试的那一期走：原因是同一个，只是现场在前一期那里。
    by = {r["form"]: r for r in out}
    for r in out:
        ev = r["evidence"]
        src = by.get(ev.get("after")) if ev.get("signal") == "not-attempted" else None
        if src is not None and src is not r:
            explain, owner, retry, action = CAUSES[src["cause"]]
            r.update(cause=src["cause"], owner=owner, retry=retry, action=action,
                     explain="%s（%s 失败后脚本没有再试这一期，原因同它）"
                             % (explain, src["form"]))
    return out


def merge_final(previous, entries=(), no_data=(), now=None):
    """补收之后的最终结论。纯函数。

    previous   这个下载目录上一次写下的缺失记录（店内诊断的结果）
    entries    登记簿里**现在**还属于这个目录的条目
    no_data    补收时平台明说没有数据的条目

    之前记为"欠着"（owed）的全部作废重算：补收取回了的自然不再出现；还欠着的
    用补收时最后一次探测的证据重新归因 —— 店内只探测过一次、只能说"分不清"
    的，这时往往已经分得清了。其余的（同步路线、申请阶段就失败的、店内就确认
    没数据的）结论不会再变，原样保留。不为了诊断重开店铺。
    """
    out = [r for r in (previous or []) if not r.get("owed")]
    seen = {(r.get("form"), r.get("period") or "") for r in out}
    for e, owed in ([(e, True) for e in entries or ()]
                    + [(e, False) for e in no_data or ()]):
        rec = record_from_entry(e, now=now)
        if owed:
            rec["owed"] = True
        key = (rec["form"], rec["period"])
        if key not in seen:
            seen.add(key)
            out.append(rec)
    return out


def evidence_ref(record):
    """给人看的"证据在哪"：优先截图旁边那份 json。"""
    ev = record.get("evidence") or {}
    return ev.get("file") or ev.get("screenshot") or ""


def format_section(rows, cautions=None, expired=None):
    """批次末尾那一段"未下载的报表与原因"。返回行列表，调用方负责打印。

    rows: [{"store", "missing": [记录...], "cautions": [...]}]
    同一家店同一个原因的报表并成一行 —— 店铺打不开时 16 张报表各打一行，
    真正要看的那一行就被淹了。
    """
    L = ["", "=" * 60, "未下载的报表与原因", "=" * 60]
    every = []
    for row in rows:
        groups = []
        for r in row.get("missing") or []:
            every.append(r)
            g = next((g for g in groups if g[0]["cause"] == r["cause"]
                      and g[0]["explain"] == r["explain"]), None)
            if g is None:
                groups.append([r])
            else:
                g.append(r)
        for g in groups:
            first = g[0]
            forms = "、".join(
                "%s%s" % (r["form"], "（%s）" % r["period"] if r.get("period") else "")
                for r in g[:6]) + (" 等 %d 张" % len(g) if len(g) > 6 else "")
            L.append("  %s | %s" % (row.get("store"), forms))
            L.append("      原因    [%s] %s" % (first["cause"], first["explain"]))
            if not any(affects_report(r) for r in g):
                L.append("      影响    财务报表不读这张，缺了不影响报表")
            L.append("      责任方  %s ｜ 重试%s"
                     % (first["owner"], "有用" if first["retry"] else "没用"))
            L.append("      下一步  %s" % first["action"])
            ref = evidence_ref(first)
            if ref:
                L.append("      证据    %s" % ref)
    for r in expired or []:
        every.append(r)
        L.append("  %s | %s%s" % (r.get("store"), r["form"],
                                  "（%s）" % r["period"] if r.get("period") else ""))
        L.append("      原因    [%s] %s" % (r["cause"], r["explain"]))
        L.append("      下一步  %s" % r["action"])
    if not every:
        L.append("  没有。应下的报表全部到齐。")
    else:
        L.append("")
        L.append("  按原因统计：")
        for code, n in sorted(count_by_cause(every).items(),
                              key=lambda kv: (-kv[1], kv[0])):
            L.append("    %-24s %3d  %s%s"
                     % (code, n, CAUSES[code][0],
                        "" if code not in NO_ATTENTION else "（无需处理）"))
    old = [(row.get("store"), r) for row in rows
           for r in (row.get("carried") or [])]
    if old:
        L.append("")
        L.append("  更早几轮留下的（不计入本轮，结论在它们各自目录的 REPORT.md 里）：")
        for store, r in old:
            L.append("    %s | %s | 来自 %s | %s"
                     % (store, r["form"], r.get("carried_from") or "?",
                        "最新一趟已下到，旧条目已撤销" if r.get("superseded")
                        else "[%s] %s" % (r["cause"], r["explain"])))
    notes = [(row.get("store"), c) for row in rows
             for c in (row.get("cautions") or [])]
    if cautions:
        notes.extend(cautions)
    if notes:
        L.append("")
        L.append("  要留意（申请后列表里没出现新行，按已有的行匹配；"
                 "若随后下到了文件，可能是更早的那份）：")
        by_store = {}
        for store, c in notes:
            by_store.setdefault(store, []).append(c)
        for store, cs in by_store.items():
            L.append("    %s：%s" % (store, "；".join(
                "%s 申请 %s → 采用 %s" % (c.get("form"), c.get("requested") or "?",
                                         c.get("adopted") or "?") for c in cs)))
    return L
