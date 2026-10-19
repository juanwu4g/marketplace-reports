# -*- coding: utf-8 -*-
"""Update this machine to the latest code, safely, in one step (update.cmd).

    python -m tools.update            # pull, install, self-check, dry run
    python -m tools.update --no-pull  # only install + checks (after a manual pull)

What it does, and why each step is there:

1. refuses while a monthly run holds logs/monthly.lock (a pull in the middle of a
   run mixes old and new code in one month's results);
2. refuses when tracked files have local edits (a pull would fail half-way or
   overwrite someone's fix; config.json is git-ignored and never touched);
3. `git pull --ff-only` (never merges, never rewrites history);
4. `pip install -r requirements.txt` with PYTHONUTF8=1 — without it pip on a
   Chinese Windows cannot read requirements.txt (gbk) and installs nothing;
5. bootstrap.py (versions, config.json, store setup, Ziniao client) and
   `run_monthly.py --region ALL --dry-run` (plans every market; opens nothing).

If a check fails after the pull, it prints the one command that returns to the
code that was there before. Exit 0 = updated and checked; 1 = something failed
(the message says what); 2 = refused before changing anything.
"""
import argparse
import os
import subprocess
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import run_monthly
from core import runlog
from core.paths import ROOT


def say(text=""):
    print(text, flush=True)


def run(cmd, env=None, quiet=False):
    """Run a command in the project folder; (exit code, combined output)."""
    try:
        p = subprocess.run(cmd, cwd=ROOT, env=env, stdout=subprocess.PIPE,
                           stderr=subprocess.STDOUT)
    except OSError as e:
        return 127, str(e)
    out = p.stdout.decode("utf-8", "replace")
    if not quiet and out.strip():
        say(out.rstrip())
    return p.returncode, out


def monthly_running():
    """Who holds the monthly lock, if a live process does."""
    import re
    try:
        with open(run_monthly.LOCK, encoding="utf-8") as f:
            who = f.read().strip()
    except (IOError, OSError):
        return None
    m = re.search(r"pid=(\d+)", who)
    if m and run_monthly.alive(int(m.group(1))):
        return who
    return None


def local_edits():
    """Tracked files changed on this machine (untracked/ignored files don't count)."""
    code, out = run(["git", "status", "--porcelain", "--untracked-files=no"], quiet=True)
    if code != 0:
        return None
    return [l for l in out.splitlines() if l.strip()]


def main(argv=None):
    ap = argparse.ArgumentParser(description="update this machine to the latest code")
    ap.add_argument("--no-pull", action="store_true", help="skip git pull")
    a = ap.parse_args(argv)
    py = sys.executable
    say("=" * 62)
    say(" 更新代码 —— %s" % ROOT)
    say("=" * 62)

    who = monthly_running()
    if who:
        say("[拒绝] 月度报表正在运行（%s）。等它跑完再更新。" % who)
        return 2

    before = runlog.git_commit(ROOT)
    if not a.no_pull:
        if not os.path.exists(os.path.join(ROOT, ".git")):
            say("[拒绝] 这个文件夹不是 git 仓库（没有 .git），没法自动更新。")
            say("       请按 README 第 2 节重新 git clone。")
            return 2
        edits = local_edits()
        if edits is None:
            say("[拒绝] 找不到 git 命令。请先安装 Git for Windows。")
            return 2
        if edits:
            say("[拒绝] 这台机器上有改过的代码文件，自动更新会和它们冲突：")
            for l in edits[:10]:
                say("       %s" % l)
            say("       请联系维护的人确认这些改动要不要保留，再更新。")
            return 2
        say("\n[1/4] 拉取最新代码（当前 %s）" % (before or "?"))
        code, out = run(["git", "pull", "--ff-only"])
        if code != 0:
            say("[失败] git pull 没成功（上面是 git 的原话）。代码没有改动。")
            return 1
    after = runlog.git_commit(ROOT)
    back = ("git checkout %s" % before) if before and before != after else None
    if before and after and before != after:
        say("       已更新：%s -> %s" % (before, after))
    else:
        say("       已经是最新：%s" % (after or "?"))

    def failed(what):
        say("\n[失败] %s" % what)
        if back:
            say("       想退回更新前的代码：在本文件夹运行  %s" % back)
            say("       （之后再运行 update.cmd 会重新拉到最新）")
        return 1

    say("\n[2/4] 安装依赖（requirements.txt）")
    env = dict(os.environ, PYTHONUTF8="1")
    code, _ = run([py, "-m", "pip", "install", "-q", "--disable-pip-version-check",
                   "-r", "requirements.txt"], env=env)
    if code != 0:
        return failed("依赖没装上（上面是 pip 的原话）")

    say("\n[3/4] 环境自检（bootstrap.py）")
    code, _ = run([py, "bootstrap.py"], env=env)
    if code != 0:
        return failed("环境自检没通过：按上面的提示处理")

    say("\n[4/4] 试运行：只列出每个国家/平台会做什么，不打开任何店铺")
    code, out = run([py, "run_monthly.py", "--region", "ALL", "--dry-run"], env=env,
                    quiet=True)
    tail = [l for l in out.splitlines() if "[dry-run]" in l or "错误" in l]
    for l in tail[-20:]:
        say("       " + l.strip())
    if code != 0:
        say(out.rstrip()[-2000:])
        return failed("试运行出错（退出码 %d）" % code)

    say("\n" + "=" * 62)
    say(" 更新完成：代码版本 %s" % runlog.code_version(ROOT))
    say("=" * 62)
    return 0


if __name__ == "__main__":
    sys.exit(main())
