# -*- coding: utf-8 -*-
"""Several run days per month (docs/RESULTS_AND_PERIODS.md §6): one
CalendarTrigger with several <Day>s; gui_settings keeps a list and still reads
the old single day; task names and arguments do not change."""
import pytest

from core import region_choice as RC
from core import win_tasks as W


def test_task_xml_with_several_days_is_one_trigger():
    x = W.task_xml(RC.task_args("ALL"), day=[26, 3, 12, 3], hour=2, minute=0, who="PC\\meli")
    assert x.count("<CalendarTrigger>") == 1
    assert "<DaysOfMonth><Day>3</Day><Day>12</Day><Day>26</Day></DaysOfMonth>" in x
    assert "<Arguments>--entrance scheduled --region ALL</Arguments>" in x      # unchanged
    one = W.task_xml(RC.task_args("ALL"), day=15, who="PC\\meli")
    assert "<DaysOfMonth><Day>15</Day></DaysOfMonth>" in one
    assert W.task_xml(RC.task_args("ALL"), day=[15], who="PC\\meli") == one
    # the rehearsal task still has no trigger
    assert "<Triggers></Triggers>" in W.task_xml(RC.rehearsal_args("ALL"), who="PC\\meli")
    for bad in ([], [0], [32], ["x"]):
        with pytest.raises(ValueError):
            W.task_xml(day=bad, who="PC\\meli")


def test_registered_days_are_read_back():
    st = W.parse_state({"EXISTS": "1", "XML": W.task_xml("--entrance scheduled --region ALL",
                                                          day=[3, 12, 26], hour=2, who="a")})
    assert st["DAYS"] == [3, 12, 26] and st["DAY"] == "3,12,26" and st["TIME"] == "02:00"
    assert RC.task_when(st) == "每月 3、12、26 号 02:00"
    old = W.parse_state({"EXISTS": "1", "XML": W.task_xml("--entrance scheduled", day=15, hour=2, who="a")})
    assert old["DAY"] == "15" and old["DAYS"] == [15]
    assert RC.task_when(old) == "每月 15 号 02:00"
    assert RC.task_when({"DAY": "15", "TIME": "02:00"}) == "每月 15 号 02:00"      # without DAYS


def test_settings_keep_a_list_and_read_old_single_days():
    assert RC.settings_days({}) == [15]
    assert RC.settings_days({"sched_day": 9}) == [9]                         # oldest key
    assert RC.settings_days({"sched_day_ALL": 20, "sched_day": 9}) == [20]
    assert RC.settings_days({"sched_days_ALL": [26, 3, 12]}) == [3, 12, 26]
    assert RC.settings_days({"sched_days_ALL": ["x"], "sched_day_ALL": 7}) == [7]   # broken list
    assert RC.settings_days({"sched_days_ALL": [40], "sched_day_ALL": 40}) == [15]
    saved = RC.days_settings([26, 3, 12])
    assert saved == {"sched_days_ALL": [3, 12, 26], "sched_day_ALL": 3}       # old windows read day 3
    assert RC.settings_days(saved) == [3, 12, 26]


def test_parse_days():
    assert RC.parse_days("3, 12，26") == [3, 12, 26]
    assert RC.parse_days(" 26 3 3 ") == [3, 26]
    assert RC.parse_days("15") == [15]
    for bad in ("", "0", "29", "3, x", "-1"):
        with pytest.raises(ValueError):
            RC.parse_days(bad)
    assert RC.days_text([3, 12]) == "3、12"


def test_clash_with_several_days():
    assert RC.clash(([3, 15], "02:00"), (15, "03:00"))
    assert "15 号" in RC.clash(([3, 15], "02:00"), {"DAYS": [15, 26], "TIME": "02:30"})
    assert RC.clash(([3, 15], "02:00"), ([4, 16], "02:00")) is None
    assert RC.clash((15, "02:00"), {"DAY": "15", "TIME": "02:10"})          # old states still work
