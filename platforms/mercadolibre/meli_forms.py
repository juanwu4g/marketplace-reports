# -*- coding: utf-8 -*-
"""
MercadoLibre seller-panel download flows, driven from the left-tab menu.

Two methods, both starting from the seller panel:

  download_billing_reports(driver, ...)
      Facturación -> latest N months -> Ir al detalle -> Reportes
      -> Seleccionar todos los reportes -> Descargar

      The number of files is NOT fixed: MercadoLibre only offers the report
      types that have data for that period, so select-all yields a different
      count per month. Observed for BOCINA_SM on 2026-08-28:
          Agosto (EN CURSO)  4  BILL_ML, FULL, PAYMENT, NC_ML
          Julio  (cerrado)   4  BILL_ML, FULL, PAYMENT, NC_ML
          Junio  (cerrado)   2  BILL_ML, PAYMENT      (only 2 types offered)
      Callers should not hard-code an expected count.

  download_sales_excel(driver, ...)
      Ventas -> remove the "Envíos de hoy" filter -> Descargar Excel de ventas
      -> wait for the notification widget -> Descargar on the newest entry.
      Yields exactly 1 file.

DOM notes that matter:
  * The Reportes panel is a micro-frontend rendered inside a SHADOW ROOT
    (section.remote-module.fbi-billing-fe-reporting-reporting). Ordinary
    querySelector / By.XPATH cannot see into it - every lookup here goes
    through the deep helpers below.
  * The select-all checkbox has a React-generated id (e.g. "«r2»") that changes
    between renders. It must be found by its label text, never by id.
  * The per-report ids ARE stable: BILL_ML, FULL, PAYMENT, NC_ML.
"""
import datetime
import os
import calendar
import time

from core import site_time
from core import download_diag
from platforms.mercadolibre import meli_site as site_config

# Rebound per store by set_site(); the values here are the Mexican defaults the
# module was written against.
BILLING_BASE = "https://vendedores.mercadolibre.com.mx"
BILLING_RESUME = BILLING_BASE + "/billing/resume"
VENTAS_LISTADO = BILLING_BASE + "/ventas/omni/listado"
SPACE_MANAGEMENT = BILLING_BASE + "/publicaciones/listado/space_management"
SELLER_HOME = BILLING_BASE + "/resumen"
WANT_LOCALE = "es_MX"


def set_site(code):
    """Point every URL and the language gate at one country site ("mx"/"cl").

    Call before a store's routes run. The paths are ASSUMED identical across
    sites - same seller centre, different host. That has not been verified
    against a live Chilean store yet; a path that turns out to differ belongs
    in meli_site.py.
    """
    global BILLING_BASE, BILLING_RESUME, VENTAS_LISTADO, SPACE_MANAGEMENT
    global SELLER_HOME, WANT_LOCALE
    facts = site_config.get(code)
    BILLING_BASE = facts["seller_host"]
    BILLING_RESUME = BILLING_BASE + "/billing/resume"
    VENTAS_LISTADO = BILLING_BASE + "/ventas/omni/listado"
    SPACE_MANAGEMENT = BILLING_BASE + "/publicaciones/listado/space_management"
    SELLER_HOME = BILLING_BASE + "/resumen"
    WANT_LOCALE = facts["locale"]
    sep = facts.get("september", "septiembre")
    ES_MONTHS[9] = sep
    SPANISH_MONTH_NAMES[8] = sep.capitalize()

# Spanish month names, as they appear in the datepicker's aria-labels.
ES_MONTHS = [None, "enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# ---------------------------------------------------------------- shadow DOM

# Every helper below recurses through open shadow roots. Returning the element
# itself lets Selenium wrap it as a WebElement, so we can issue a real click
# (React handles synthetic events from real clicks more reliably than from
# dispatchEvent).

JS_DEEP_QUERY = """
const sel = arguments[0];
const scan = (root) => {
  const hit = root.querySelector(sel);
  if (hit) return hit;
  for (const el of root.querySelectorAll('*')) {
    if (el.shadowRoot) { const r = scan(el.shadowRoot); if (r) return r; }
  }
  return null;
};
return scan(document);
"""

JS_DEEP_CHECKBOX_BY_LABEL = """
const want = arguments[0].toLowerCase();
const scan = (root) => {
  for (const el of root.querySelectorAll('input[type=checkbox]')) {
    if ((el.id||'').startsWith('nav-')) continue;
    const holder = el.closest('label,div,li,tr');
    if (((holder && holder.innerText) || '').toLowerCase().includes(want)) return el;
  }
  for (const el of root.querySelectorAll('*')) {
    if (el.shadowRoot) { const r = scan(el.shadowRoot); if (r) return r; }
  }
  return null;
};
return scan(document);
"""

JS_DEEP_CHECKBOX_STATE = """
const scan = (root, out) => {
  root.querySelectorAll('input[type=checkbox]').forEach(e => {
    if ((e.id||'').startsWith('nav-')) return;
    out.push({id: e.id||'', checked: e.checked,
              label: (e.closest('label,div')?.innerText||'').trim()
                     .replace(/\\s+/g,' ').slice(0,50)});
  });
  root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) scan(el.shadowRoot, out); });
  return out;
};
return scan(document, []);
"""

JS_WIDGET_ROWS = """
return [...document.querySelectorAll('.process-notification-process')].map(e => ({
  title: (e.querySelector('.process-notification-process__content-title')?.innerText||'').trim(),
  sub:   (e.querySelector('.process-notification-process__content-subtitle')?.innerText||'').trim(),
  date:  (e.querySelector('.process-notification-process__content-date')?.innerText||'').trim(),
  hasDl: !!e.querySelector('.process-notification-link')
}));
"""


def deep_find(driver, css):
    return driver.execute_script(JS_DEEP_QUERY, css)


def deep_checkbox(driver, label_text):
    return driver.execute_script(JS_DEEP_CHECKBOX_BY_LABEL, label_text)


def click(driver, el):
    """Scroll into view, then real-click, falling back to a JS click."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.6)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


# ---------------------------------------------------------------- downloads

def snapshot_dir(path):
    if not os.path.isdir(path):
        return {}
    # _diag 是诊断证据的子目录，不是下载。不跳过的话，等待下载的途中它被建
    # 出来，就会被当成"新到的文件"。
    return {f: os.path.getsize(os.path.join(path, f)) for f in os.listdir(path)
            if f != download_diag.DIAG_DIR}


def checked_report_ids(state, select_all_id=None):
    """JS_DEEP_CHECKBOX_STATE 的结果 -> 已勾选的报表 id，去重、保持顺序。

    那段 JS 把每个 root 里的复选框都收一遍，同一个框会被收到两次：日志里
    4 张报表显示"已勾选 8"，min_files=8 永远凑不齐，于是每一期都要等满
    settle*3 而不是 settle，还次次报"勾选了 8 张，实际得到 4 个文件"。
    """
    return download_diag.dedupe(
        c["id"] for c in (state or [])
        if c.get("checked") and c.get("id") and c["id"] != select_all_id)


def set_download_dir(driver, path):
    """Route downloads to `path` for this session via CDP.

    Without this, files land in the store's own downloadPath, which is shared
    across runs and makes counting new files unreliable.
    """
    if not os.path.isdir(path):
        os.makedirs(path)
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow", "downloadPath": path})
        return True
    except Exception as e:
        print("  [警告] 无法重定向下载目录：%s" % str(e)[:90])
        return False


PARTIAL_SUFFIXES = (".crdownload", ".tmp")

# 下载目录 -> 上一次等待超时时还没写完的文件（去掉 .crdownload / .tmp 的名字）。
# 它们迟早会被 Chrome 改名成完整文件，那时出现在**下一条路线**的等待里，会被
# 当成下一条路线的文件。见 _late_names。
LATE_PARTIALS = {}


def _partial_stem(name):
    for suf in PARTIAL_SUFFIXES:
        if name.endswith(suf):
            return name[:-len(suf)]
    return None


def _late_names(path, before):
    """这次等待不能算作"新文件"的完整文件名。

    快照里一个文件还是 export.xlsx.crdownload，Chrome 写完后把它改名成
    export.xlsx —— 名字不在快照里，按"快照里没有就是新的"会被当成这一次点击
    下到的文件。于是上一条路线迟到的文件被下一条路线收走（库存的文件记到了
    账单名下之类），而下一条路线自己的文件反倒被当成多余的。

    排除两类：快照里已经在写的（.crdownload / .tmp 去掉后缀的名字），以及
    之前某次等待超时时还在写的（LATE_PARTIALS）。只认 Chrome 这种"原名 +
    后缀"的命名；"Unconfirmed 12345.crdownload" 那种对不上最终文件名，管不到。
    """
    late = set(LATE_PARTIALS.get(os.path.abspath(path), ()))
    late.update(s for s in (_partial_stem(n) for n in before or ()) if s)
    return late


def wait_for_downloads(path, before, min_files=1, timeout=240, settle=12, label="",
                       start_timeout=None):
    """Block until downloads have *settled*, and return the new filenames.

    "Select all reports" fires several downloads a few seconds apart, so
    returning at the first completed file loses the rest. There is no signal
    from the site saying "that was the last one", so completion is inferred
    from the download folder going quiet:

      - got everything we expected (`min_files`) and quiet for `settle`  -> return
      - fewer than expected, but quiet for `settle * 3`                  -> give up

    The second branch matters: `min_files` is a prediction (the number of report
    checkboxes ticked), not a promise. Without the escape hatch a wrong
    prediction would burn the whole `timeout` on every affected period.

    `start_timeout`：多少秒内**连一个下载都没开始**（连 .crdownload 都没有）
    就放弃。上面那个逃生口要求至少见过一个文件，所以点了之后什么都没发生时它
    永远不触发，会干等满 `timeout`。9/21 凌晨平台的库存报表没有响应，
    costos×2 + general 每家店白等 3×240 秒，把一轮 3 小时的批次拖到 5 小时
    以上，被计划任务的超时整批杀掉，报表那一步根本没跑到。默认 None 保持原
    行为 —— 只给"点了就该立刻开始下载"的路线用。
    """
    deadline = time.time() + timeout
    began = time.time()
    seen = set()
    last_change = time.time()
    late = _late_names(path, before)
    key = os.path.abspath(path)
    pending = []
    while time.time() < deadline:
        time.sleep(2)
        new = set(snapshot_dir(path)) - set(before)
        landed = new & late
        if landed:
            # 上一条路线迟到的文件，不是这一次点击下到的
            new -= landed
            left = LATE_PARTIALS.get(key)
            if left:
                left.difference_update(landed)
        if start_timeout and not new and time.time() - began >= start_timeout:
            print("  [警告] %s：%d 秒内没有任何下载开始，平台可能没响应，放弃等待"
                  % (label or "download", start_timeout))
            return []
        pending = [k for k in new if k.endswith(PARTIAL_SUFFIXES)]
        done = {k for k in new if not k.endswith(PARTIAL_SUFFIXES)}
        if done != seen:
            seen = done
            last_change = time.time()
        quiet = time.time() - last_change
        if seen and not pending and (
                (len(seen) >= min_files and quiet >= settle)
                or quiet >= settle * 3):
            return sorted(seen)
    if pending:
        # 超时了还有没写完的：记下来，它写完时别让后面的等待把它认领走
        LATE_PARTIALS.setdefault(key, set()).update(
            _partial_stem(k) for k in pending)
    if len(seen) < min_files:
        print("  [警告] %s：预期至少 %d 个文件，等了 %d 秒实际只有 %d 个"
              % (label or "download", min_files, timeout, len(seen)))
    return sorted(seen)


# ---------------------------------------------------- 0. 语言关卡

# 整个下载层靠可见的西语文字定位元素（"Ir al detalle"、"descargar"、
# "almacenamiento"、/ultimos?\s*(\d+)\s*mes/ …），导出文件的列名也是西语，
# 而**导出文件的语言跟着账号的界面语言走**。有人在浏览器里把语言改成 English，
# 后果分两种，后一种才可怕：
#
#   * 元素找不到 -> 抛错，这家店整个失败，看得见
#   * 账期页的月份名核对失败 -> **那一期被静默跳过**，报表少算费用而毫无提示。
#     只下目标月而漏掉后续账期，实测让 EWTTO_SM 少算佣金 973.23、代扣税率
#     从 9.172% 虚高到 9.221%。
#
# 与其把几十处选择器翻译成多语言（还要为每种语言维护一套列名表），不如在跑
# 任何路线之前先确认语言，不对就改回来。
#
# 关键点：判据用的是 `data-selected-locale="es_MX"` 这个 **locale 代码**，
# 不是可见文字。所以这道检查本身不会因为语言变了而失效 —— 用文字判断语言，
# 语言一变判断自己就先瞎了。
JS_GET_LOCALE = """
const el = document.querySelector('.nav-user-menu-language-switcher');
if (el && el.getAttribute('data-selected-locale'))
  return el.getAttribute('data-selected-locale');
const s = document.querySelector(
  '.nav-user-menu-language-switcher__option[aria-selected="true"]');
return s ? s.getAttribute('data-value') : null;
"""

# 切换器藏在用户菜单里（折叠状态），但 JS 点击不受可见性限制：先点触发器把
# 组件唤醒，再点目标选项。已经是目标语言时直接返回 already，不做无谓点击。
JS_SET_LOCALE = """
const want = arguments[0];
const root = document.querySelector('.nav-user-menu-language-switcher');
if (!root) return 'no-switcher';
if (root.getAttribute('data-selected-locale') === want) return 'already';
const t = document.getElementById('nav-user-menu-language-switcher-trigger');
if (t) t.click();
const o = root.querySelector('[data-value="' + want + '"]');
if (!o) return 'no-option';
o.click();
return 'clicked';
"""


def ensure_language(driver, want=None, timeout=30):
    """进店后、跑任何路线之前：确认卖家后台是西语，不是就切回去。

    返回 (进来时的 locale 或 None, 是否动过)。

    读不到切换器时**不拦路**，只发警告：那多半是页面改版，而为了一个读不到
    的判据把 12 家店全挡住，比它要防的问题更糟。真的改不回来才抛错 —— 那种
    情况下继续跑必然产出错语言的导出。
    """
    want = want or WANT_LOCALE      # 调用时取值：set_site() 之后才是对的
    if not (driver.current_url or "").startswith(BILLING_BASE):
        driver.get(SELLER_HOME)
        time.sleep(6)
    cur = driver.execute_script(JS_GET_LOCALE)
    if cur is None:
        print("  [警告] 顶栏找不到语言切换器，无法确认界面语言。"
              "若平台改版，下载层的西语选择器可能已经失效。")
        return None, False
    if cur == want:
        print("  界面语言：%s ✓" % cur)
        return cur, False

    print("  [警告] 界面语言是 %s，不是 %s —— 导出文件的列名会跟着变，"
          "正在切回去。" % (cur, want))
    r = driver.execute_script(JS_SET_LOCALE, want)
    if r not in ("clicked", "already"):
        raise RuntimeError("切换界面语言失败（%s）：当前 %s，需要 %s。"
                           "请在浏览器里手动改回西班牙语。" % (r, cur, want))
    time.sleep(4)
    driver.get(SELLER_HOME)
    time.sleep(6)
    now = driver.execute_script(JS_GET_LOCALE)
    if now != want:
        raise RuntimeError("界面语言切换后仍是 %s（期望 %s）。"
                           "请在浏览器里手动改回西班牙语再跑。" % (now, want))
    print("  界面语言已从 %s 改回 %s ✓" % (cur, want))
    return cur, True


# ---------------------------------------------------- 1. Facturación reports

def _open_reports_tab(driver, timeout=90):
    """Select the Reportes tab and wait for its shadow content to render."""
    tab = deep_find(driver, "button.billing-detail_tab-reports")
    if tab is None:
        raise RuntimeError("Reportes tab not found")
    if "andes-tab--selected" not in (tab.get_attribute("class") or ""):
        click(driver, tab)

    deadline = time.time() + timeout
    while time.time() < deadline:
        boxes = driver.execute_script(JS_DEEP_CHECKBOX_STATE)
        if boxes:
            return boxes
        time.sleep(2)
    raise RuntimeError("Reportes panel never rendered")


JS_CARD_INFO = """
let e = arguments[0];
for (let i = 0; i < 8 && e; i++) {
  e = e.parentElement; if (!e) break;
  const txt = e.innerText || '';
  const m = txt.match(
    /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Sep?tiembre|Octubre|Noviembre|Diciembre)/i);
  if (m) return {month: m[1], enCurso: /EN CURSO/i.test(txt)};
}
return {month: '?', enCurso: false};
"""


def billing_periods(driver, months=2, skip_current=False):
    """Return the indices of the 'Ir al detalle' buttons to visit, newest first.

    `skip_current` drops the EN CURSO period - its charges are still accruing,
    so its reports are provisional.
    """
    driver.get(BILLING_RESUME)
    time.sleep(8)
    buttons = driver.find_elements("xpath", "//button[normalize-space()='Ir al detalle']")
    picked = []
    for i, b in enumerate(buttons):
        info = driver.execute_script(JS_CARD_INFO, b)
        tag = "EN CURSO" if info["enCurso"] else "cerrado"
        if skip_current and info["enCurso"]:
            print("  跳过 %s（%s）" % (info["month"], tag))
            continue
        picked.append((i, info["month"]))
        if len(picked) >= months:
            break
    return picked


def download_billing_reports(driver, download_dir, months=2, timeout=240,
                             skip_current=False, diag=None):
    """Facturación -> newest `months` periods -> all reports -> Descargar.

    Returns {period: [filenames]}.

    `diag`（可选）是这家店的 download_diag.Collector。常规路线事先不知道账期
    的月份名，所以报表 id 按新旧次序编号：billing/1 是最新一期。
    """
    results = {}
    picked = billing_periods(driver, months=months, skip_current=skip_current)
    print("  本次下载的账期：%s" % ", ".join(m for _, m in picked))
    # 要 N 期、页面只给得出更少：缺的那几期得有个说法。一期都没有多半是页面
    # 没渲染出来或被顶掉了，和"这家店只有一期账单"不是一回事。
    for slot in range(len(picked) + 1, months + 1):
        if picked:
            download_diag.give_up(
                diag, driver, "billing/%d" % slot, "not-offered", shot=False,
                step="resume 页只有 %d 个可下的账期" % len(picked))
        else:
            download_diag.give_up(
                diag, driver, "billing/%d" % slot, "ui-missing",
                step="resume 页上没有任何 'Ir al detalle' 按钮")

    for slot, (idx, _month) in enumerate(picked, 1):
        form = "billing/%d" % slot
        driver.get(BILLING_RESUME)
        time.sleep(8)

        buttons = driver.find_elements(
            "xpath", "//button[normalize-space()='Ir al detalle']")
        if idx >= len(buttons):
            print("  [警告] 只有 %d 个账期可选" % len(buttons))
            for later in range(slot, len(picked) + 1):
                download_diag.give_up(
                    diag, driver, "billing/%d" % later, "ui-missing",
                    step="重新打开 resume 页后只剩 %d 个 'Ir al detalle'，"
                         "要点第 %d 个" % (len(buttons), idx + 1),
                    buttons=len(buttons))
            break

        month = driver.execute_script("""
            let e = arguments[0];
            for (let i = 0; i < 8 && e; i++) {
              e = e.parentElement; if (!e) break;
              const m = (e.innerText||'').match(
                /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Sep?tiembre|Octubre|Noviembre|Diciembre)/i);
              if (m) return m[1];
            }
            return '?';
        """, buttons[idx])

        print("\n--- Facturación [%d] %s：进入明细（Ir al detalle）---" % (idx, month))
        click(driver, buttons[idx])
        time.sleep(9)

        period = driver.current_url.rstrip("/").split("/")[-1].split("?")[0]
        print("    明细页：%s" % driver.current_url)

        key = "%s (%s)" % (month, period)
        try:
            files = _download_detail_reports(driver, download_dir, month, timeout,
                                             diag=diag, form=form, period=key)
        except RuntimeError as e:
            if diag is None:
                raise                    # 老调用方：行为不变，交给 _guard
            # 一期失败不该连累其余各期，更不该让已经下好的那几期从结果里
            # 消失 —— 以前异常一路冒到 _guard，整个 billing 结果变成 {}，
            # 文件明明在目录里，汇总却说一张都没有。
            print("    [警告] %s：%s" % (key, e))
            download_diag.give_up(diag, driver, form, "exception", period=key,
                                  exception_class=e.__class__.__name__,
                                  exception=str(e), step="明细页 Reportes 面板")
            files = []
        results[key] = files

    return results


SPANISH_MONTH_NAMES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                       "Julio", "Agosto", "Septiembre", "Octubre",
                       "Noviembre", "Diciembre"]


def billing_detail_url(year, month):
    """某个会计月的账单明细页地址。

    明细页按**账期结束日**寻址，而结束日就是当月最后一天 —— 实测
    Agosto → 20260831、Septiembre → 20260930、Julio → 20260731 都成立。

    这条路径的意义：resume 页只摆最近 3 期卡片，更早的账期在页面上根本
    点不到，但明细页本身还在，直接访问就能拿到。做历史月份全靠它。
    """
    last = calendar.monthrange(year, month)[1]
    return "%s/billing/detail/%04d%02d%02d" % (BILLING_BASE, year, month, last)


def download_billing_for_months(driver, download_dir, year, month,
                                extra=2, timeout=240, diag=None):
    """下载目标会计月**以及随后 extra 期**的账单。

    为什么不能只下目标月：订单级费用按**计费日**入账，月末几天的订单，佣金常常
    落在下一期账单上。实测 EWTTO_SM 的 7 月订单有 56 行费用记在 8/9 月账单里
    ——只下 7 月的话纯佣金少算 973.23，而代扣税是用「Ventas 合并值 − 纯佣金」
    倒推的，于是等额虚高，税率从 9.172% 变成 9.221%。

    **还没开始的账期直接不去请求。** 原来的假设是"超出当月的期数会被页面校验
    挡掉，不必自己算上限"——挡是挡得住，但不便宜：明细页打开后要等 Reportes
    面板 90 秒超时，才报「Reportes panel never rendered」。9 月中旬跑
    --month 2026-08 时，8 月后面的第 2 期是 10 月，每家店都为它白等一次、
    报一条吓人的警告，12 家店就是约 18 分钟。当月（还没结账）的账期仍然要下
    —— 它的明细页已经存在，月末订单的费用就记在里面。
    """
    out = {}
    y, mo = year, month
    today = site_time.site_today()
    n = max(1, extra + 1)
    for i in range(n):
        if datetime.date(y, mo, 1) > today:
            print("    跳过 %04d-%02d 的账单：账期还没开始" % (y, mo))
            break
        form = "billing/%04d-%02d" % (y, mo)
        # 只有目标月是"必须有"的：它拿不到，这一趟就白跑了。
        got = download_billing_for_period(driver, download_dir, y, mo, timeout,
                                          diag=diag, form=form, required=(i == 0))
        if not got and i > 0:
            # 后续期拿不到很正常（还没到/已关闭），不算失败。但再往后那几期
            # 是因为这里停了才没去试的，得记一笔，否则它们看起来像凭空少了。
            yy, mm = y, mo
            for _ in range(i + 1, n):
                mm += 1
                if mm > 12:
                    yy, mm = yy + 1, 1
                if datetime.date(yy, mm, 1) > today:
                    break
                download_diag.give_up(diag, None, "billing/%04d-%02d" % (yy, mm),
                                      "not-attempted", after=form, shot=False)
            break
        out.update(got)
        mo += 1
        if mo > 12:
            y, mo = y + 1, 1
    return out


def download_billing_for_period(driver, download_dir, year, month, timeout=240,
                                diag=None, form=None, required=True):
    """下载**指定会计月**的全部账单报表，不管它还在不在 resume 页的卡片里。

    返回 {期间: [文件名]}；该账期不存在或页面打不开时返回 {}（不抛异常 ——
    历史月份取不到是常态，不该让整个店铺的下载失败）。
    """
    url = billing_detail_url(year, month)
    label = "%s %d" % (SPANISH_MONTH_NAMES[month - 1], year)
    print("\n--- Facturación %s：直接访问明细页 ---" % label)
    print("    %s" % url)
    driver.get(url)
    time.sleep(9)

    # 页面存在≠是我们要的那一期。用页面自己写的月份名核对，避免平台把
    # 无效日期悄悄重定向到最近一期、结果把 9 月的账单当成 7 月的存下来。
    shown = driver.execute_script("""
        const re = /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Sep?tiembre|Octubre|Noviembre|Diciembre)/i;
        const m = (document.body.innerText || '').match(re);
        return m ? m[1] : null;
    """)
    want = SPANISH_MONTH_NAMES[month - 1]
    form = form or "billing/%04d-%02d" % (year, month)
    if shown and shown.lower() != want.lower():
        print("    [警告] 页面显示的是 %s，不是 %s —— 该账期可能已不可访问，跳过"
              % (shown, want))
        download_diag.give_up(diag, driver, form, "wrong-period-page",
                              period=label, required=required,
                              shown=shown, want=want, asked_url=url)
        return {}
    if not shown:
        print("    [警告] 页面没有可识别的月份，跳过")
        download_diag.give_up(diag, driver, form, "no-month-on-page",
                              period=label, required=required, asked_url=url)
        return {}

    period = url.rstrip("/").split("/")[-1]
    key = "%s (%s)" % (want, period)
    try:
        files = _download_detail_reports(driver, download_dir, want, timeout,
                                         diag=diag, form=form, period=key,
                                         required=required)
    except RuntimeError as e:
        print("    [警告] %s：%s" % (label, e))
        download_diag.give_up(diag, driver, form, "exception", period=label,
                              required=required,
                              exception_class=e.__class__.__name__,
                              exception=str(e), step="明细页 Reportes 面板")
        return {}
    return {key: files}


def _download_detail_reports(driver, download_dir, month, timeout=240,
                             diag=None, form=None, period=None, required=None):
    """已经站在某个账期的明细页上：勾选全部报表并下载。返回文件名列表。

    从 download_billing_reports 里抽出来，好让"按指定账期下载"能走同一段代码 ——
    两条路径的差别只在怎么到达明细页，页面上的操作完全一样。
    """
    _open_reports_tab(driver)

    # 勾"全选"要确认真的勾上了。2026-10-14 TRQWH_TA01 的 9 月账单：页面、
    # 标签页都对，点了全选却一个都没勾上（截图里四个框全空），Descargar 一直
    # 是灰的，这一期账单就没下到。同一家店 10 月那一期紧接着就正常 —— 是点得
    # 太早、表格还没就绪，不是页面变了。所以没勾上就重新找、重新点。
    checked, select_all_id = [], None
    for attempt in range(3):
        cb = deep_checkbox(driver, "seleccionar todos los reportes")
        if cb is None:
            raise RuntimeError("select-all checkbox not found for %s" % month)
        # The select-all id is React-generated ("«r2»") and changes between
        # renders, so remember this element's own id rather than pattern
        # matching it out of the list below.
        select_all_id = cb.get_attribute("id")
        if not cb.is_selected():
            click(driver, cb)
            time.sleep(2)
        state = driver.execute_script(JS_DEEP_CHECKBOX_STATE)
        checked = checked_report_ids(state, select_all_id)
        if checked:
            break
        print("    全选没有生效（第 %d 次），等一下再点" % (attempt + 1))
        if cb.is_selected():
            click(driver, cb)          # 框是勾上的、报表却没跟着勾：先取消再来
        time.sleep(4)
    expected = len(checked)
    print("    已勾选报表（%d）：%s" % (expected, checked or "（无）"))

    btn = None
    deadline = time.time() + 30
    while time.time() < deadline:
        btn = deep_find(driver, "button.billing-reporting-download__button")
        if btn is not None and btn.is_enabled():
            break
        time.sleep(1)
    if btn is None or not btn.is_enabled():
        raise RuntimeError("bulk Descargar never enabled for %s" % month)

    before = snapshot_dir(download_dir)
    click(driver, btn)
    print("    已点击 Descargar，等待文件落盘…")
    files = wait_for_downloads(download_dir, before,
                               min_files=max(expected, 1),
                               timeout=timeout, label="billing %s" % month)
    for f in files:
        print("      + %s" % f)
    form = form or "billing/%s" % month
    if not files:
        # 点了 Descargar，一个文件都没到
        download_diag.give_up(
            diag, driver, form, "download-not-landed", period=period,
            required=required, checked=checked,
            **download_diag.folder_delta(before, snapshot_dir(download_dir)))
        return files
    download_diag.delivered(diag, form, files, period)
    if expected and len(files) != expected:
        _got, missing, unknown = download_diag.match_billing_files(checked, files)
        print("    [警告] %s：勾选了 %d 张报表，实际得到 %d 个文件%s"
              % (month, expected, len(files),
                 "，缺的是：%s" % "、".join(missing) if missing else ""))
        # 这一期算到了，缺的那一张单独记：billing/<期>/<报表 id>
        if len(files) < expected:
            for cid in (missing or unknown or ["unidentified"]):
                download_diag.give_up(
                    diag, driver, "%s/%s" % (form, cid), "short-download",
                    period=period, checked=checked, files=files,
                    missing_reports=missing, unrecognised_ids=unknown)
    return files


# --------------------------------------------------------- 2. Ventas Excel

# Accent-insensitive text match. The UI writes "Últimos 6 meses"; a plain
# lowercase compare fails on the accented U, so both sides are normalised.
JS_NORM = """
const norm = (s) => (s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                    .toLowerCase().replace(/\\s+/g,' ').trim();
"""

JS_FIND_PERIOD_TRIGGER = """
// Verified live: the control is a combobox button inside .sc-dropdown-date-range
return document.querySelector('.sc-dropdown-date-range button.andes-dropdown__trigger')
    || document.querySelector('button.andes-dropdown__trigger');
"""

# NEVER match an option by its full text. The menu renders the label followed by
# the concrete date range it resolves to today:
#     "Últimos 2 meses  29 jun. al 29 ago."
#     "Últimos 6 meses  28 feb. al 29 ago."
# Those dates move every single day, so an exact compare works once and then
# silently stops finding the option. Match the label only.
JS_FIND_PERIOD_OPTION = JS_NORM + """
const re = new RegExp('ultimos?\\\\s*' + arguments[0] + '\\\\s*mes');
for (const el of document.querySelectorAll(
        'li[role=option], [role=option], .andes-list__item')) {
  if (re.test(norm(el.innerText))) return el;
}
return null;
"""

JS_LIST_PERIOD_OPTIONS = """
const out = [];
for (const el of document.querySelectorAll('li[role=option], [role=option]')) {
  const t = (el.innerText||'').trim().replace(/\\s+/g,' ');
  if (t && t.length < 60) out.push(t);
}
return [...new Set(out)];
"""

# The trigger's innerText repeats the label (a visually-hidden copy sits beside
# the visible one, giving "Últimos 2 meses Últimos 2 meses"), so read the
# display span instead of the button.
JS_CURRENT_PERIOD = JS_NORM + """
const el = document.querySelector(
      '.sc-dropdown-date-range .andes-dropdown__display-values')
   || document.querySelector('.andes-dropdown__display-values');
if (!el) return null;
const m = norm(el.innerText).match(/ultimos?\\s*(\\d+)\\s*mes/);
return m ? parseInt(m[1], 10) : null;
"""


def set_sales_period(driver, months=6, timeout=45):
    """Ventas -> click the period dropdown -> pick "Últimos N meses".

    The human flow is: click "Últimos 2 meses", a menu appears, click
    "Últimos 6 meses". There is no URL parameter for this that survives a
    reload, so it has to be driven through the menu.

    Returns True if the period reads N months afterwards.
    """
    want_label = "Últimos %d meses" % months

    current = driver.execute_script(JS_CURRENT_PERIOD)
    if current == months:
        print("    时间范围已是 %s" % want_label)
        return True
    print("    当前时间范围为 %s 个月，切换为 %d 个月" % (current or "?", months))

    trigger = driver.execute_script(JS_FIND_PERIOD_TRIGGER)
    if trigger is None:
        print("    [警告] 未找到时间范围下拉框，保持原值不变")
        return False
    click(driver, trigger)
    time.sleep(2)

    option = None
    deadline = time.time() + 15
    while time.time() < deadline:
        option = driver.execute_script(JS_FIND_PERIOD_OPTION, months)
        if option is not None:
            break
        time.sleep(1)
    if option is None:
        print("    [警告] 菜单中没有 %r。可选项为：%s"
              % (want_label, driver.execute_script(JS_LIST_PERIOD_OPTIONS)))
        return False
    print("    选择：%s" % option.text.replace("\n", " ").strip()[:52])

    click(driver, option)

    # the list reloads; wait for the trigger to actually read the new period
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        if driver.execute_script(JS_CURRENT_PERIOD) == months:
            print("    时间范围已设为 %s" % want_label)
            time.sleep(20)         # let the order list finish reloading
            return True
    print("    [警告] 时间范围未能切换到 %s" % want_label)
    return False


# 最近一次 request_sales_excel 实际生效的时间范围。用模块级变量而不是改返回值，
# 是因为返回值的形状被 run_downloads 的 pending 机制约束着，改它要牵动收取逻辑。
SALES_PERIOD_ACTUAL = {"months": None}


def request_sales_excel(driver, period_months=6, diag=None):
    """PHASE 1 - ask MercadoLibre to build the sales Excel, do not wait.

    Generation is server-side (~25-30s) and the file appears in the corner
    notification widget when done, so the wait can overlap with other routes.
    Returns a pending dict, or None if the button never became available.
    """
    print("\n--- Ventas：申请 Excel de ventas ---")
    driver.get(VENTAS_LISTADO)
    time.sleep(12)

    # remove the "Envíos de hoy" applied filter; without this the list is
    # today's shipments only (often 0 sales) and the Excel button stays disabled
    removed = driver.execute_script("""
        for (const t of document.querySelectorAll('[data-testid="filters--applied"] .andes-tag')) {
          const lbl = (t.querySelector('.andes-tag__label')?.innerText||'').trim();
          if (/env[ií]os de hoy/i.test(lbl)) {
            const b = t.querySelector('button.andes-tag__close');
            if (b) { b.click(); return lbl; }
          }
        }
        return null;
    """)
    time.sleep(10)
    today_filter = driver.execute_script("""
        return [...document.querySelectorAll('[data-testid="filters--applied"] .andes-tag')]
          .some(t => /env[ií]os de hoy/i.test(
              (t.querySelector('.andes-tag__label')?.innerText||'').trim()));
    """)
    if today_filter:
        print("    [警告] Envíos de hoy 筛选仍在，Ventas 按钮可能保持禁用")
    else:
        print("    Envíos de hoy 筛选：%s" % ("已移除" if removed else "未发现已应用"))

    # widen the period AFTER clearing the filter - the dropdown re-renders when
    # filters change, and a handle grabbed before that goes stale
    # 读回**实际**生效的窗口，而不是我们请求的那个。菜单里没有的档位会被
    # set_sales_period 拒绝并保持原值（实测要 3 拿到的是 2），而这一步失败
    # 之后导出照样能跑，只是少了一大段数据 —— 不读回来就完全看不出。
    if period_months:
        set_sales_period(driver, months=period_months)
    actual = driver.execute_script(JS_CURRENT_PERIOD)
    print("    时间范围核验：请求 %s 个月，实际 %s 个月" % (period_months, actual))
    if period_months and actual != period_months:
        print("    [警告] 请求 %s 个月，实际生效 %s 个月 —— 导出会少一段数据"
              % (period_months, actual))
        download_diag.flag(diag, "range_clamped", {
            "ventas": {"requested_months": period_months,
                       "actual_months": actual}})
    SALES_PERIOD_ACTUAL["months"] = actual

    btn = None
    deadline = time.time() + 180
    while time.time() < deadline:
        btn = driver.execute_script("""
            // MercadoLibre uses different classes across store accounts. Match
            // the action's label, then prefer a visible, enabled button.
            return [...document.querySelectorAll('button')]
              .find(e => /Descargar Excel de ventas/i.test(
                  (e.getAttribute('aria-label')||'') + ' ' + (e.innerText||''))
                  && e.getClientRects().length > 0
                  && !e.matches(':disabled')
                  && e.getAttribute('aria-disabled') !== 'true') || null;
        """)
        if btn is not None:
            break
        time.sleep(2)
    if btn is None:
        candidates = driver.execute_script("""
            return [...document.querySelectorAll('button')]
              .filter(e => /Descargar Excel de ventas/i.test(
                  (e.getAttribute('aria-label')||'') + ' ' + (e.innerText||'')))
              .map(e => ({visible: e.getClientRects().length > 0,
                          disabled: e.matches(':disabled') ||
                                    e.getAttribute('aria-disabled') === 'true'}));
        """)
        reason = ("未找到按钮" if not candidates else
                  "找到 %d 个按钮，但都不可点击：%s" % (len(candidates), candidates))
        print("    [警告] 'Descargar Excel de ventas' 在 180 秒内不可点击（%s）" % reason)
        download_diag.give_up(
            diag, driver, "ventas",
            "ui-disabled" if candidates else "ui-missing",
            step="'Descargar Excel de ventas' 在 180 秒内不可点击（%s）"
                 % ("按钮在但被禁用" if candidates else "没有这个按钮"),
            candidates=candidates, today_filter_still_applied=bool(today_filter),
            period_months_actual=actual)
        return None

    rows_before = len(driver.execute_script(JS_WIDGET_ROWS))
    click(driver, btn)
    print("    已提交申请，文件在服务端继续生成")
    return {"kind": "ventas", "rows_before": rows_before,
            "requested_at": time.time()}


def collect_sales_excel(driver, pending, download_dir, timeout=600, poll=15,
                        diag=None):
    """PHASE 3 - return to Ventas and download the generated Excel.

    Polls up to `timeout` seconds printing status. Not-ready is reported, not
    raised: the file stays in the widget and can be collected on a later run.
    """
    waited = time.time() - pending.get("requested_at", time.time())
    print("\n--- Ventas：取回 Excel de ventas ---")
    print("    距提交申请已过去 %.0f 秒" % waited)

    if "ventas/omni" not in driver.current_url:
        driver.get(VENTAS_LISTADO)
        time.sleep(12)

    before = snapshot_dir(download_dir)
    rows_before = pending.get("rows_before", 0)

    ready = False
    deadline = time.time() + timeout
    while time.time() < deadline:
        rows = driver.execute_script(JS_WIDGET_ROWS)
        left = int(deadline - time.time())
        if len(rows) > rows_before and rows[0]["hasDl"] and "xito" in rows[0]["sub"]:
            print("    [剩余 %4d 秒] 就绪：%s | %s"
                  % (left, rows[0]["date"], rows[0]["sub"]))
            ready = True
            break
        top = rows[0]["sub"] if rows else "(no entries)"
        print("    [剩余 %4d 秒] %s" % (left, top))
        time.sleep(poll)

    if not ready:
        print("    [警告] %d 秒内没有新的成功记录，改用最新一条"
              % timeout)

    clicked = driver.execute_script("""
        const row = document.querySelector('.process-notification-process');
        if (!row) return null;
        const link = row.querySelector('.process-notification-link');
        if (!link) return null;
        link.scrollIntoView({block:'center'});
        link.click();
        return (row.querySelector('.process-notification-process__content-date')?.innerText||'').trim();
    """)
    if clicked is None:
        print("    [警告] 最新一条通知里没有 Descargar 链接")
        download_diag.give_up(diag, driver, "ventas", "ui-missing",
                              step="最新一条通知里没有 Descargar 链接",
                              ready=ready)
        return []
    print("    已在第 %s 条记录上点击 Descargar" % clicked)

    files = wait_for_downloads(download_dir, before, min_files=1,
                               timeout=180, settle=8, label="ventas excel")
    for f in files:
        print("      + %s" % f)
    if not files:
        download_diag.give_up(
            diag, driver, "ventas", "download-not-landed",
            **download_diag.folder_delta(before, snapshot_dir(download_dir)))
    return files


def download_sales_excel(driver, download_dir, timeout=300, period_months=6,
                         diag=None):
    """Request and collect in one call - for running the Ventas route alone.

    The split version above is what a full run uses, so the generation wait
    overlaps with the other routes.
    """
    pending = request_sales_excel(driver, period_months=period_months, diag=diag)
    if not pending:
        return []
    return collect_sales_excel(driver, pending, download_dir, timeout=timeout,
                               diag=diag)



# ------------------------------------------------- 3. Stock reports (Full)

# Publicaciones -> Control de stock -> "Descargar reportes de stock" -> 4 reports.
#
# Verified live 2026-08-29. Each behaves differently:
#
#   a. Reporte de costos por el servicio de almacenamiento
#         a "Periodo" dropdown of billing periods; pick the newest, then Descargar
#   b. Reporte general de stock
#         NO period control - it is a point-in-time snapshot ("Actualizado el
#         29 de agosto"), so a date range would be meaningless. Downloads on click.
#   c. Reporte consolidado de movimientos
#         a range datepicker, then Aplicar, then Descargar
#   d. Reporte de devoluciones a la bodega
#         downloads on click
#
# Three mechanics that are easy to get wrong:
#
#   * The "Descargar reportes de stock" dropdown opens ONLY via a JS click. A
#     native Selenium click and an ActionChains click both leave
#     aria-expanded="false". (The billing Reportes tab is the opposite - it
#     needs a real click. Do not "harmonise" these.)
#   * Each menu <li> is inert; the real target is an empty overlay
#     button.andes-list__item-actionable inside it.
#   * In the range picker, pick the EARLIER date first. Once a start is
#     selected every earlier day becomes --disabled, so doing it backwards
#     silently selects only one end of the range.

# 库存报表点了就该立刻开始下载（9/17 那一轮每家店都是几秒内落盘）。
# 60 秒内连个 .crdownload 都没出现，就是平台没响应，不必等满 240 秒。
STOCK_START_TIMEOUT = 60

STOCK_REPORTS = [
    ("costos",       "reporte de costos por el servicio de almacenamiento"),
    ("general",      "reporte general de stock"),
    ("consolidado",  "reporte consolidado de movimientos"),
    # 墨西哥叫 "a la bodega"，智利叫 "al centro de distribución"（2026-10-15 实测）。
    # 用 "|" 隔开的几种叫法依次试，哪个点得到算哪个。
    ("devoluciones", "reporte de devoluciones a la bodega|"
                     "reporte de devoluciones al centro de distribución"),
]

JS_OPEN_STOCK_MENU = JS_NORM + """
for (const el of document.querySelectorAll('button')) {
  if (/descargar reportes de stock/.test(norm(el.innerText))) {
    el.click();               // JS click only - see the note above
    return true;
  }
}
return false;
"""

# MercadoLibre changed this menu's markup: the entries used to be
# li.andes-button-dropdown__menu-item and are now plain li.andes-list__item.
# Accept either, or the whole stock route silently returns "menu has: []" while
# the menu is in fact open (aria-expanded="true") right in front of it.
JS_STOCK_MENU_SEL = ("li.andes-button-dropdown__menu-item, "
                     "li.andes-list__item")

JS_STOCK_ITEMS = """
return [...document.querySelectorAll('%s')]
       .map(e => (e.innerText||'').trim())
       .filter(t => /reporte/i.test(t));
""" % JS_STOCK_MENU_SEL

JS_CLICK_STOCK_ITEM = JS_NORM + """
const want = norm(arguments[0]);
for (const li of document.querySelectorAll('%s')) {
  if (norm(li.innerText).indexOf(want) === 0) {
    (li.querySelector('button.andes-list__item-actionable') || li).click();
    return true;
  }
}
return false;
""" % JS_STOCK_MENU_SEL

# Find the modal by its heading text, so we never grab the page-level
# "Descargar reportes de stock" button sitting behind the overlay.
JS_MODAL_BUTTON = """
const key = arguments[0].toLowerCase();
for (const el of document.querySelectorAll('[role=dialog], .andes-modal')) {
  if (!(el.innerText||'').toLowerCase().includes(key)) continue;
  for (const b of el.querySelectorAll('button')) {
    if (/descargar/i.test(b.innerText||'')) return {found: true, disabled: b.disabled};
  }
}
return {found: false, disabled: true};
"""

JS_CLICK_MODAL_BUTTON = """
const key = arguments[0].toLowerCase();
for (const el of document.querySelectorAll('[role=dialog], .andes-modal')) {
  if (!(el.innerText||'').toLowerCase().includes(key)) continue;
  for (const b of el.querySelectorAll('button')) {
    if (/descargar/i.test(b.innerText||'')) { b.click(); return true; }
  }
}
return false;
"""

JS_OPEN_PERIOD_DROPDOWN = """
for (const el of document.querySelectorAll('[role=dialog],.andes-modal')) {
  if (!/almacenamiento/i.test(el.innerText||'')) continue;
  const t = el.querySelector('button.andes-dropdown__trigger');
  if (t) { t.click(); return true; }
}
return false;
"""

# Billing-period options read like "Del 1 de julio al 31 de julio del 2026".
# The Espanol / English / Chinese entries in the DOM belong to the nav language
# switcher, so match the date shape rather than taking the first list item.
# Two shapes appear in this dropdown, and which ones are present varies:
#   "Período actual - ..."                     the open, still-accruing period
#   "Del 1 de julio al 31 de julio del 2026"   a closed period
# The list is already newest-first, so DOM order is preserved and the caller
# just takes the first N. On 2026-08-31 only the two closed periods were
# offered - no "Período actual" - so nothing may assume it exists.
JS_LIST_PERIODS = """
const re = /^(per[ií]odo\\s+actual|del\\s+\\d+\\s+de\\s+\\w+)/i;
const out = [];
for (const el of document.querySelectorAll('li[role=option], li.andes-list__item')) {
  const t = (el.innerText||'').trim().replace(/\\s+/g, ' ');
  if (re.test(t)) out.push(t);
}
return [...new Set(out)];
"""

# Pick by text rather than by index: the option list is rebuilt each time the
# modal opens, and an index would silently drift if the newest period rolls
# over between iterations.
#
# Whitespace MUST be collapsed on both sides. JS_LIST_PERIODS normalises, and
# the "Período actual" entry is multi-line ("... del 2026\\nAcumulas $ 415.81
# hasta el momento"), so comparing against raw innerText never matches it - the
# current period is silently skipped while the closed ones still work.
JS_PICK_PERIOD = """
const flat = (s) => (s||'').trim().replace(/\\s+/g, ' ');
const want = flat(arguments[0]);
for (const el of document.querySelectorAll('li[role=option], li.andes-list__item')) {
  if (flat(el.innerText) === want) {
    (el.querySelector('button.andes-list__item-actionable') || el).click();
    return true;
  }
}
return false;
"""

JS_CAL_CAPTION = """
const e = document.querySelector('.andes-datepicker__caption-label');
return e ? e.innerText.trim() : null;
"""

JS_CAL_NAV = """
const b = document.querySelector(arguments[0]);
if (!b) return false;
b.click();
return true;
"""

# aria-label is "sabado 29 de agosto de 2026 " - endsWith on the trimmed label
# is exact and needs no regex escaping.
JS_CAL_PICK = """
const want = arguments[0] + ' de ' + arguments[1] + ' de ' + arguments[2];
for (const b of document.querySelectorAll('button.andes-datepicker__day')) {
  if ((b.getAttribute('aria-label')||'').trim().endsWith(want)) {
    const cell = b.closest('.andes-datepicker__cell');
    if (cell && /--disabled/.test(cell.className)) return 'disabled';
    b.click();
    return 'ok';
  }
}
return 'not-found';
"""

JS_CAL_FIRST_ENABLED = """
for (const b of document.querySelectorAll('button.andes-datepicker__day')) {
  const cell = b.closest('.andes-datepicker__cell');
  if (cell && /--disabled/.test(cell.className)) continue;
  return parseInt((b.innerText||'').trim(), 10);
}
return null;
"""

JS_OPEN_DATEPICKER = """
const t = document.querySelector('#report-date-picker__trigger');
if (!t) return false;
t.click();
return true;
"""

JS_CLICK_APLICAR = """
for (const b of document.querySelectorAll('button')) {
  if (b.getBoundingClientRect().width > 4 &&
      /^aplicar$/i.test((b.innerText||'').trim())) { b.click(); return true; }
}
return false;
"""


def _cal_goto(driver, year, month, tries=18):
    """Walk the single-month calendar to year/month."""
    want = "%s %d" % (ES_MONTHS[month], year)
    for _ in range(tries):
        cap = (driver.execute_script(JS_CAL_CAPTION) or "").lower()
        if cap == want:
            return True
        cur_m = next((i for i, n in enumerate(ES_MONTHS) if n and cap.startswith(n)), None)
        parts = cap.split()
        cur_y = int(parts[-1]) if parts and parts[-1].isdigit() else year
        back = (cur_y, cur_m or 0) > (year, month)
        driver.execute_script(
            JS_CAL_NAV,
            ".andes-datepicker__button--previous" if back
            else ".andes-datepicker__button--next")
        time.sleep(1.2)
    return False


def _select_date_range(driver, months_back=2, diag=None):
    """Select (today - months_back) .. today in the range picker.

    The picker enforces its own floor - on 2026-08-29 every day before 1 July
    was disabled - so a requested start that is out of range is clamped forward
    to the earliest day the calendar actually offers, rather than failing or
    silently selecting only one end of the range.
    """
    today = site_time.site_today()
    m, y = today.month - months_back, today.year
    while m <= 0:
        m += 12
        y -= 1
    start_day = min(today.day, 28)

    if not driver.execute_script(JS_OPEN_DATEPICKER):
        print("      [警告] 未找到日期选择器")
        return False
    time.sleep(3)

    picked_start = None
    for bump in range(months_back + 1):
        mm, yy = m + bump, y
        while mm > 12:
            mm -= 12
            yy += 1
        if not _cal_goto(driver, yy, mm):
            continue

        # Only the originally requested month keeps the requested day. Once we
        # have had to skip forward, the target day is arbitrary - taking it
        # would throw away the front of the month (June disabled + day 28 gave
        # 28 Jul instead of 1 Jul, halving the range), so take the earliest day
        # the calendar allows.
        if bump == 0:
            if driver.execute_script(JS_CAL_PICK, start_day,
                                     ES_MONTHS[mm], yy) == "ok":
                picked_start = "%04d-%02d-%02d" % (yy, mm, start_day)
                break
        first = driver.execute_script(JS_CAL_FIRST_ENABLED)
        if first and driver.execute_script(
                JS_CAL_PICK, first, ES_MONTHS[mm], yy) == "ok":
            picked_start = "%04d-%02d-%02d" % (yy, mm, first)
            print("      起始日期收敛为 %s（更早的日期不可选）" % picked_start)
            download_diag.flag(diag, "range_clamped", {
                "stock/consolidado": {
                    "requested_start": "%04d-%02d-%02d" % (y, m, start_day),
                    "actual_start": picked_start}})
            break
    if not picked_start:
        print("      [警告] 没有可选的起始日期")
        return False
    time.sleep(1)

    _cal_goto(driver, today.year, today.month)
    res = driver.execute_script(JS_CAL_PICK, today.day,
                                ES_MONTHS[today.month], today.year)
    if res != "ok":
        print("      [警告] 结束日期 %s -> %s" % (today, res))
    print("      区间：%s .. %s" % (picked_start, today))
    time.sleep(1)

    driver.execute_script(JS_CLICK_APLICAR)
    time.sleep(3)
    return True



def _open_stock_menu(driver, timeout=60):
    """Open the stock-reports dropdown, waiting for the page to be ready.

    A fixed sleep is not enough: this page finishes rendering late, and under
    load a 12s wait left the button absent, so the whole route returned
    "menu has: []" and downloaded nothing. Poll for the menu ITEMS, not just
    the trigger - the trigger can exist while the list is still empty.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if driver.execute_script(JS_OPEN_STOCK_MENU):
            time.sleep(3)
            if driver.execute_script(JS_STOCK_ITEMS):
                return True
        time.sleep(3)
    return False


def _download_storage_costs(driver, download_dir, label, periods=2, timeout=240,
                            diag=None):
    """Report (a), once per billing period, newest first.

    The modal only holds one period at a time, so this reopens it per period
    rather than trying to multi-select.
    """
    out = []
    known = None
    for i in range(periods):
        driver.get(SPACE_MANAGEMENT)
        time.sleep(12)
        before = snapshot_dir(download_dir)

        form = "stock/costos/%d" % (i + 1)
        rest = ["stock/costos/%d" % (j + 1) for j in range(i, periods)]
        if not _open_stock_menu(driver):
            print("      [警告] 库存报表菜单始终为空")
            for f in rest:           # break：这一期和之后的都不会再试
                download_diag.give_up(diag, driver, f, "ui-missing",
                                      step="库存报表菜单始终为空", menu_items=[])
            break
        if not click_stock_item(driver, label):
            print("      [警告] 未找到该菜单项")
            if diag is not None:
                items = driver.execute_script(JS_STOCK_ITEMS)
                for f in rest:
                    download_diag.give_up(diag, driver, f, "ui-missing",
                                          step="菜单里没有 %r" % label,
                                          menu_items=items)
            break
        time.sleep(5)

        driver.execute_script(JS_OPEN_PERIOD_DROPDOWN)
        time.sleep(3)
        available = driver.execute_script(JS_LIST_PERIODS)
        if known is None:
            known = available
            print("      可选账期：%d 个 -> %s" % (len(known), known))
        if i >= len(available):
            print("      仅有 %d 个账期可选，停止" % len(available))
            for f in rest:
                if available:
                    # 平台就只给这么多期，不是谁的错
                    download_diag.give_up(
                        diag, driver, f, "not-offered", shot=False,
                        step="仓储费只有 %d 个账期可选" % len(available),
                        available=available)
                else:
                    # 一期都没有：分不清是下拉框没打开还是这家店没有仓储费，
                    # 留现场给人看
                    download_diag.give_up(
                        diag, driver, f, "ui-missing",
                        step="仓储费弹窗的 Periodo 下拉框里没有任何账期",
                        available=available)
            break

        want = available[i]
        if not driver.execute_script(JS_PICK_PERIOD, want):
            print("      [警告] 无法选中 %r" % want)
            download_diag.give_up(diag, driver, form, "ui-missing", period=want,
                                  step="无法选中账期 %r" % want,
                                  available=available)
            continue
        print("      [%d/%d] %s" % (i + 1, min(periods, len(available)), want))
        time.sleep(2)

        st = driver.execute_script(JS_MODAL_BUTTON, "almacenamiento")
        if not st["found"] or st["disabled"]:
            print("      [警告] 该账期的 Descargar 不可用")
            download_diag.give_up(
                diag, driver, form,
                "ui-disabled" if st["found"] else "ui-missing", period=want,
                step="仓储费弹窗的 Descargar 不可用", modal_button=st)
            continue
        driver.execute_script(JS_CLICK_MODAL_BUTTON, "almacenamiento")

        files = wait_for_downloads(download_dir, before, min_files=1,
                                   timeout=timeout, settle=6,
                                   label="stock/costos[%d]" % i,
                                   start_timeout=STOCK_START_TIMEOUT)
        for f in files:
            print("        + %s" % f)
        out.extend(files)
        if files:
            download_diag.delivered(diag, form, files, want)
        else:
            download_diag.give_up(
                diag, driver, form, "download-not-landed", period=want,
                step="已点击弹窗的 Descargar",
                **download_diag.folder_delta(before, snapshot_dir(download_dir)))
    return out


def click_stock_item(driver, label):
    """点库存报表菜单里的一项。label 可以是用 "|" 隔开的几种叫法（不同国家站点）。"""
    for one in label.split("|"):
        if driver.execute_script(JS_CLICK_STOCK_ITEM, one):
            return True
    return False


# 这几项点一下菜单就直接下载（没有弹窗要填）：没有下载开始时值得重点一次。
STOCK_RETRY_ON_NO_START = ("general", "devoluciones")


def download_stock_reports(driver, download_dir, months_back=2, timeout=240,
                           only=None, storage_periods=2, diag=None):
    """Download the four stock reports. Returns {key: [filenames]}.

    `storage_periods` - how many billing periods to pull for report (a),
    newest first. 2 gives the newest and the one below it.
    """
    print("\n--- Publicaciones：库存报表（reportes de stock）---")
    results = {}

    for key, label in STOCK_REPORTS:
        if only and key not in only:
            continue

        if key == "costos":
            print("    [costos] %s" % label)
            results[key] = _download_storage_costs(
                driver, download_dir, label,
                periods=storage_periods, timeout=timeout, diag=diag)
            continue

        form = "stock/%s" % key

        driver.get(SPACE_MANAGEMENT)
        time.sleep(12)
        before = snapshot_dir(download_dir)

        if not _open_stock_menu(driver):
            print("    [警告] 库存报表菜单始终为空，跳过 %s" % key)
            download_diag.give_up(diag, driver, form, "ui-missing",
                                  step="库存报表菜单始终为空", menu_items=[])
            results[key] = []
            continue

        if not click_stock_item(driver, label):
            items = driver.execute_script(JS_STOCK_ITEMS)
            print("    [警告] 未找到菜单项 %r，当前菜单为：%s" % (label, items))
            download_diag.give_up(diag, driver, form, "ui-missing",
                                  step="菜单里没有 %r" % label, menu_items=items)
            results[key] = []
            continue
        print("    [%s] %s" % (key, label))
        time.sleep(5)

        # 没有文件时该怪谁：下面任何一步没成，之后"没有下载开始"只是它的后果。
        why = None
        if key == "consolidado":
            if not _select_date_range(driver, months_back=months_back, diag=diag):
                why = ("ui-missing", "日期区间没有选上")
            # only (c) still needs its modal's Descargar; (b) and (d) download
            # on click, and (a) is handled by _download_storage_costs above
            st = driver.execute_script(JS_MODAL_BUTTON, "consolidado")
            if not st["found"]:
                print("      [警告] 弹窗里没有 Descargar 按钮")
                why = why or ("ui-missing", "弹窗里没有 Descargar 按钮")
            elif st["disabled"]:
                print("      [警告] Descargar 仍不可点击，账期未被接受")
                why = why or ("ui-disabled",
                              "弹窗的 Descargar 仍不可点击，账期未被接受")
            else:
                driver.execute_script(JS_CLICK_MODAL_BUTTON, "consolidado")

        files = wait_for_downloads(download_dir, before, min_files=1,
                                   timeout=timeout, settle=6,
                                   label="stock/%s" % key,
                                   start_timeout=STOCK_START_TIMEOUT)
        if not files and not why and key in STOCK_RETRY_ON_NO_START:
            # 点了菜单项，一个下载都没开始。2026-10-14 那轮库存总表有 3 家店这样
            # （TRQWH_TA01、EWXON_EW、DAREU_TA，前后半小时内），同一批店 10-09
            # 都下到了 —— 平台一时没响应，不是页面变了。重开页面再点一次；
            # `before` 不变，第一次点击的文件晚到也照样算数。
            print("      没有下载开始，重开页面再点一次")
            driver.get(SPACE_MANAGEMENT)
            time.sleep(12)
            if _open_stock_menu(driver) and click_stock_item(driver, label):
                time.sleep(5)
                files = wait_for_downloads(download_dir, before, min_files=1,
                                           timeout=timeout, settle=6,
                                           label="stock/%s（重试）" % key,
                                           start_timeout=STOCK_START_TIMEOUT)
        for f in files:
            print("      + %s" % f)
        results[key] = files
        if files:
            download_diag.delivered(diag, form, files)
        else:
            download_diag.give_up(
                diag, driver, form, why[0] if why else "download-not-landed",
                step=why[1] if why else "点击菜单项后没有下载开始",
                **download_diag.folder_delta(before, snapshot_dir(download_dir)))

    return results


def probe_sales_excel(driver, pending, settle=10):
    """Cheap readiness check for the Ventas Excel. Leaves the driver on the page."""
    if "ventas/omni" not in driver.current_url:
        driver.get(VENTAS_LISTADO)
        time.sleep(settle)
    else:
        driver.refresh()
        time.sleep(settle)
    rows = driver.execute_script(JS_WIDGET_ROWS)
    # 这一眼看到了什么，留给诊断用。写在条目上而不是模块变量里 —— 条目是
    # 这家店、这张报表自己的。带下划线：调用方读完就弹掉，不进 pending.json。
    pending["_probe"] = {
        "row_state": "ventas-widget", "widget_rows": len(rows or []),
        "rows_before": pending.get("rows_before", 0),
        "top_sub": rows[0]["sub"] if rows else "",
        "top_has_dl": bool(rows and rows[0]["hasDl"])}
    return bool(rows
                and len(rows) > pending.get("rows_before", 0)
                and rows[0]["hasDl"]
                and "xito" in rows[0]["sub"])


def fetch_sales_excel(driver, pending, download_dir):
    """Download the generated Excel. Assumes probe_sales_excel() just passed."""
    before = snapshot_dir(download_dir)
    clicked = driver.execute_script("""
        const row = document.querySelector('.process-notification-process');
        if (!row) return null;
        const link = row.querySelector('.process-notification-link');
        if (!link) return null;
        link.scrollIntoView({block:'center'});
        link.click();
        return (row.querySelector('.process-notification-process__content-date')?.innerText||'').trim();
    """)
    if clicked is None:
        print("      [警告] 最新一条通知里没有 Descargar 链接")
        pending["_fetch"] = "no-click-target"
        return []
    pending["_fetch"] = "clicked"
    files = wait_for_downloads(download_dir, before, min_files=1,
                               timeout=180, settle=8, label="ventas excel")
    for f in files:
        print("      + %s" % f)
    return files
