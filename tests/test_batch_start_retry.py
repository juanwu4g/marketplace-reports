# -*- coding: utf-8 -*-
"""紫鸟在内核下载期间回 -10000：等一等再试，别让这家店直接失败。"""
import pytest

from core import ziniao_client
from core.ziniao_client import ZiniaoClient, ZiniaoError


def _client(responses, monkeypatch):
    c = object.__new__(ZiniaoClient)
    calls, slept = [], []
    def send(data, timeout=120):
        calls.append(data)
        return responses.pop(0)
    c.send = send
    monkeypatch.setattr(ziniao_client.time, "sleep", slept.append)
    return c, calls, slept


UNKNOWN = {"statusCode": -10000, "err": "发生未知错误"}


def test_unknown_error_is_retried_until_the_store_opens(monkeypatch, capsys):
    c, calls, slept = _client([UNKNOWN, UNKNOWN, {"statusCode": 0, "debuggingPort": 1}],
                              monkeypatch)
    info = c._start_browser({"action": "startBrowser", "browserOauth": "x"})
    assert info["debuggingPort"] == 1
    assert len(calls) == 3 and slept == [15, 30]
    assert "内核还在下载" in capsys.readouterr().out


def test_gives_up_after_the_last_wait_with_the_original_error(monkeypatch):
    n = len(ZiniaoClient.START_RETRY_WAITS) + 1
    c, calls, slept = _client([dict(UNKNOWN) for _ in range(n)], monkeypatch)
    with pytest.raises(ZiniaoError) as e:
        c._start_browser({"action": "startBrowser"})
    assert len(calls) == n and slept == list(ZiniaoClient.START_RETRY_WAITS)
    assert "-10000" in str(e.value) and e.value.ziniao_step == "startBrowser"


@pytest.mark.parametrize("resp", [{"statusCode": -10003, "err": "login"},
                                  {"statusCode": -1}, None])
def test_other_errors_are_not_retried(resp, monkeypatch):
    c, calls, slept = _client([resp], monkeypatch)
    with pytest.raises(ZiniaoError):
        c._start_browser({"action": "startBrowser"})
    assert len(calls) == 1 and slept == []
