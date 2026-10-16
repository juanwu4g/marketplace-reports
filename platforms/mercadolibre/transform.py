# -*- coding: utf-8 -*-
"""
Hand the freshly downloaded files to the cleaning pipeline in `downloads/`.

`downloads/mx_sales` is a separate project (its own README, its own pyproject,
its own tests) that cleans the raw exports into Parquet and DuckDB views. This
module is the seam between the two: it runs that project's CLI and reports what
happened. It deliberately knows nothing about how the pipeline works.

Two decisions worth keeping:

* **A separate process, not an import.** A pandas or duckdb failure must not be
  able to reach back into a batch that has just spent half an hour driving a
  browser. The worst a broken transform can do here is print an error.

* **After ALL stores, not after each one.** The pipeline is tree-shaped - its
  own README says "a snapshot tree holds the same report many times over", and
  `discover` walks the whole raw directory. Running it once at the end sees
  every store from the batch; running it per store would do the same work N
  times over.
"""
import json
import os
import subprocess
import sys
import time

from core.paths import ROOT

# The pipeline's project root: `python -m mx_sales` is run from here, so the
# package resolves without needing to be pip-installed.
PIPELINE_ROOT = os.path.join(ROOT, "downloads")
PIPELINE_PKG = os.path.join(PIPELINE_ROOT, "mx_sales")

# build 打印的那一行机器可读结果（见 downloads/mx_sales/__main__.py 的
# BUILD_RESULT_MARKER，两边必须一致）。给程序看的，不回显给人。
BUILD_RESULT_MARKER = "[BUILD_RESULT_JSON] "


def available():
    """Is the cleaning pipeline present in this checkout?"""
    return os.path.isdir(PIPELINE_PKG)


def _echo(out):
    """Print the pipeline's output, collapsing repeated lines.

    Reading the exports emits the same openpyxl "Workbook contains no default
    style" warning once per file, which buries the three lines that actually say
    what was built. Collapsed rather than filtered: a warning that fires 22 times
    is still shown, once, with its count - suppressing it outright would also
    hide the next warning nobody has seen yet.
    """
    seen = {}
    order = []
    for line in out.rstrip().splitlines():
        line = line.rstrip()
        if not line or line.startswith(BUILD_RESULT_MARKER):
            continue
        if line not in seen:
            seen[line] = 0
            order.append(line)
        seen[line] += 1
    for line in order:
        n = seen[line]
        print("    %s%s" % (line, "   (x%d)" % n if n > 1 else ""))


def _build_detail(out):
    """从 build 的输出里取出那一行机器可读结果。没有或读不懂时返回 None。"""
    detail = None
    for line in out.splitlines():
        if line.startswith(BUILD_RESULT_MARKER):
            try:
                detail = json.loads(line[len(BUILD_RESULT_MARKER):])
            except (ValueError, TypeError):
                detail = None
    return detail if isinstance(detail, dict) else None


def _invoke(subcmd, label, timeout, echo, current_runs=None):
    """Run one `python -m mx_sales <subcmd>` and describe what happened.

    Returns a dict: ok / partial / skipped / returncode / seconds / output,
    plus, for build: files_ok / files_failed / failed_current / failed_older /
    older_only (see build()).
    `partial` means some files failed and the rest were written - a state the
    caller must not report as "failed".
    Never raises - the downloads are already on disk and safe either way.
    """
    extra = {"files_ok": None, "files_failed": 0, "failed_current": [],
             "failed_older": [], "older_only": False}
    if not available():
        return dict({"ok": True, "partial": False,
                     "skipped": "downloads/ 下没有 mx_sales 包",
                     "returncode": None, "seconds": 0, "output": ""}, **extra)

    # sys.executable, so the transform runs in the same venv as the downloader.
    # pyarrow and duckdb are pinned in requirements.txt for exactly this.
    cmd = [sys.executable, "-m", "mx_sales"] + list(subcmd)
    env = dict(os.environ)
    # The pipeline prints Spanish report names; without this a pipe on Windows
    # raises UnicodeEncodeError and the whole step fails for a print statement.
    env["PYTHONIOENCODING"] = "utf-8"

    print("\n>>> %s：%s" % (label, " ".join(cmd[1:])))
    print("    工作目录：%s" % PIPELINE_ROOT)
    began = time.time()
    try:
        p = subprocess.run(cmd, cwd=PIPELINE_ROOT, env=env, timeout=timeout,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           encoding="utf-8", errors="replace")
        out, rc = p.stdout or "", p.returncode
    except subprocess.TimeoutExpired:
        return dict({"ok": False, "partial": False, "skipped": None,
                     "returncode": None, "seconds": int(time.time() - began),
                     "output": "超时（%d 秒）" % timeout}, **extra)
    except Exception as e:
        return dict({"ok": False, "partial": False, "skipped": None,
                     "returncode": None, "seconds": int(time.time() - began),
                     "output": "%s: %s" % (e.__class__.__name__, e)}, **extra)

    secs = int(time.time() - began)
    if echo and out.strip():
        _echo(out)
    # build 的退出码是三档的（见 downloads/mx_sales/__main__.py 的说明）：
    # 0 全部成功 / 1 部分文件失败但其余已写入 / 2 什么都没产出。以前这里只分
    # 零与非零，于是 55 个文件里坏 1 个也被一路报成"清洗失败"，读的人会以为
    # 整批数据都没清洗。validate 的 1 是别的意思，所以只认 build。
    is_build = bool(subcmd) and subcmd[0] == "build"
    partial = rc == 1 and is_build
    ok = rc == 0

    if is_build:
        detail = _build_detail(out)
        if detail is not None:
            failed = [f for f in detail.get("failed", []) if isinstance(f, dict)]
            extra["files_ok"] = detail.get("written")
            extra["files_failed"] = len(failed)
            extra["failed_current"] = [f for f in failed if f.get("current", True)]
            extra["failed_older"] = [f for f in failed if not f.get("current", True)]
        # build 每次都把整棵下载树重扫一遍，所以一个洗不动的旧文件会让之后
        # **每一次**清洗都退出 1，整批任务跟着每个月都"需要关注"，而本次下
        # 载的东西其实一点问题没有。调用方说清了哪些目录是本次的（current_runs），
        # 且失败的全在别的目录里，这次清洗就算成功 —— 旧文件另记在
        # failed_older 里，不丢，但不再冒充本次的失败。
        # 读不到那行机器可读结果时不做这个判断：分不清就按原来的算。
        if (partial and current_runs is not None and detail is not None
                and extra["failed_older"] and not extra["failed_current"]):
            extra["older_only"] = True
            ok, partial = True, False

    if extra["older_only"]:
        state = "成功（另有 %d 个以前的文件仍洗不动，不属于本次，见上面的 note）" \
                % len(extra["failed_older"])
    elif ok:
        state = "成功"
    elif partial:
        state = "部分文件失败（其余已写入）"
        if extra["failed_older"]:
            state += "：本次 %d 个，以前的 %d 个" % (
                len(extra["failed_current"]), len(extra["failed_older"]))
    else:
        state = "失败（退出码 %d）" % rc
    print("    %s %s，用时 %d 分 %02d 秒" % (label, state, secs // 60, secs % 60))
    return dict({"ok": ok, "partial": partial, "skipped": None,
                 "returncode": rc, "seconds": secs, "output": out[-4000:]}, **extra)


def build(args=(), timeout=3600, echo=True, current_runs=None):
    """Clean every download in the tree to Parquet and refresh the DuckDB views.

    `current_runs`：本次下载用的运行目录名（如 "20261013_090000_m202609"，给
    完整路径也行，只取最后一段），可以是一个字符串或一组。给了之后，结果里
    把洗不动的文件分成两堆：

        failed_current   本次目录里的 —— 这才是这次要管的
        failed_older     别的目录里的 —— 老问题，每次重扫都会再撞上
        older_only       True 表示失败的**全是**老文件；此时 ok=True、partial=False

    每一项是 {"report", "file"（相对 downloads/ 的路径）, "error", "current"}。
    另有 files_ok / files_failed 两个计数。

    不给（默认 None）就和以前完全一样：所有失败都算本次的，ok / partial /
    returncode 的含义不变，failed_older 恒为空。处理哪些文件也不变 —— 这个
    参数只影响"怎么报"，不影响"洗什么"。
    """
    cmd = ["build"] + list(args)
    runs = None
    if current_runs is not None:
        if isinstance(current_runs, str):
            current_runs = [current_runs]
        runs = [os.path.basename(os.path.normpath(str(r))) for r in current_runs if r]
        for r in runs:
            cmd += ["--current-run", r]
    return _invoke(cmd, "清洗", timeout, echo, current_runs=runs)


def validate(export, excel_out, timeout=1800, echo=True):
    """Apply the accounting rules to ONE ventas export and write its workbook.

    `export` must be the path RELATIVE TO downloads/, e.g.
    "UNIT_TA04/20260903_164329/20260903_Ventas_MX_...xlsx". The pipeline's
    matcher also accepts a bare filename, but that matches any file of that name
    anywhere in the tree - fine for ventas exports, whose names carry a timestamp
    and seller id, and quietly wrong for anything else. Passing the full relative
    path removes the question.

    Note this reads the RAW export, not the Parquet, and pairs it with the
    facturacion files in the SAME run folder - so the workbook is consistent to
    one snapshot and does not depend on build() having run.
    """
    excel_out = os.path.abspath(excel_out)
    # `validate --excel` writes where it is told and does not create the folder.
    parent = os.path.dirname(excel_out)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    return _invoke(["validate", export, "--excel", excel_out],
                   "对账簿", timeout, echo)
