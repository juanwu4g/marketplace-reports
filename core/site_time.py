# -*- coding: utf-8 -*-
"""平台那边的"今天"。

跑这套脚本的电脑用的是中国时间（UTC+8），MercadoLibre 墨西哥站用墨西哥城时间
（UTC-6，2022 年起不再实行夏令时）。两边差 14 小时：中国时间 0 点到下午 2 点之间，
本机的"今天"在平台那边还是"昨天"。

用本机日期算出来的结束日期于是在平台上还不存在 —— 日期选择器里那一格是灰的，
整段区间选不上，或者被悄悄收敛一天。2026-10-14 那轮 12 家店全部碰上（"结束日期
2026-10-14 -> disabled"），TANKE_EE 和 UNIT_PW01(catalog) 的 Retiros 因此申请失败。
定时任务默认凌晨 2 点跑，**每一次**都会撞上。

凡是"按今天推算平台上的日期"的地方（申请区间、哪些账期已开始、默认做哪个月），
都用这里的 site_today()；只有给本机文件起名的时间戳才继续用本机时间。
"""
import datetime

SITE_TZ = "America/Mexico_City"     # 当前站点；run_store 进店前用 set_site() 切换
# zoneinfo 读不到时区库时的退路（Windows 自带的 Python 没有 tzdata，requirements
# 里装了 tzdata；万一没装，墨西哥城现在全年都是 UTC-6）。
SITE_UTC_OFFSET_HOURS = -6


def set_site(tz, utc_offset):
    """切到另一个国家站点的时区（智利 America/Santiago，有夏令时）。"""
    global SITE_TZ, SITE_UTC_OFFSET_HOURS
    SITE_TZ, SITE_UTC_OFFSET_HOURS = tz, utc_offset


def site_now(now=None):
    """平台所在地的当前时间（不带时区的 datetime）。now 给测试用：一个带时区的时刻。"""
    now = now or datetime.datetime.now(datetime.timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        return now.astimezone(ZoneInfo(SITE_TZ)).replace(tzinfo=None)
    except Exception:
        utc = now.astimezone(datetime.timezone.utc).replace(tzinfo=None)
        return utc + datetime.timedelta(hours=SITE_UTC_OFFSET_HOURS)


def site_today(now=None):
    """平台那边的今天（date）。"""
    return site_now(now).date()
