"""사설 CA 허용(`allow_private_certificate`)이 실제로 TLS 설정에 반영되는지
검증한다. connection_security.py의 순수 로직과, bridge.py가 wss:// 에만
ssl 컨텍스트를 넘기고 ws:// 에는 안 넘기는지(websockets가 ws:// + ssl 조합을
에러로 거절하므로 중요), 그리고 compression=None을 항상 넘기는지(실제
XGEN dev 서버 상대 실측에서 확인된 회귀 방지 - 압축 확장을 켠 채 접속하면
게이트웨이가 hello 직후 close 프레임 없이 연결을 끊었다)."""
from __future__ import annotations

import ssl
from typing import Any

import pytest

from xgen_seepage.connector import bridge as bridgemod
from xgen_seepage.connector.connection_security import httpx_verify, relaxed_ssl_context


def test_relaxed_ssl_context_disables_verification() -> None:
    ctx = relaxed_ssl_context()
    assert ctx.check_hostname is False
    assert ctx.verify_mode == ssl.CERT_NONE


def test_httpx_verify_true_by_default() -> None:
    assert httpx_verify(False) is True


def test_httpx_verify_returns_relaxed_context_when_enabled() -> None:
    v = httpx_verify(True)
    assert isinstance(v, ssl.SSLContext)
    assert v.verify_mode == ssl.CERT_NONE


class _StopTest(Exception):
    """websockets.connect 를 가로챈 뒤 더 진행하지 않게 끊는 용도."""


class _FakeConnect:
    """`async with websockets.connect(...) as ws:` 형태를 흉내내는 가짜.

    실제 연결은 하지 않고, 호출 kwargs만 기록한 뒤 즉시 예외로 빠져나간다.
    """

    def __init__(self, url: str, **kwargs: Any) -> None:
        self.url = url
        self.kwargs = kwargs

    async def __aenter__(self) -> Any:
        raise _StopTest(self.kwargs)

    async def __aexit__(self, *exc: Any) -> None:
        return None


@pytest.mark.asyncio
async def test_ws_scheme_never_receives_ssl_kwarg(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, Any] = {}

    def fake_connect(url: str, **kwargs: Any) -> _FakeConnect:
        captured["url"] = url
        captured["kwargs"] = kwargs
        return _FakeConnect(url, **kwargs)

    monkeypatch.setattr(bridgemod.websockets, "connect", fake_connect)

    b = bridgemod.ConnectorMcpBridge(
        server_name="s", tool_definitions=[], call_tool=lambda n, a: {},
        get_token=lambda: _async_none(), allow_private_certificate=True,
    )
    with pytest.raises(_StopTest):
        await b._connect_once("http://plain.example.com", "u1")

    assert captured["url"].startswith("ws://")
    assert "ssl" not in captured["kwargs"]


@pytest.mark.asyncio
async def test_wss_scheme_receives_relaxed_ssl_when_enabled(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, Any] = {}

    def fake_connect(url: str, **kwargs: Any) -> _FakeConnect:
        captured["url"] = url
        captured["kwargs"] = kwargs
        return _FakeConnect(url, **kwargs)

    monkeypatch.setattr(bridgemod.websockets, "connect", fake_connect)

    b = bridgemod.ConnectorMcpBridge(
        server_name="s", tool_definitions=[], call_tool=lambda n, a: {},
        get_token=lambda: _async_none(), allow_private_certificate=True,
    )
    with pytest.raises(_StopTest):
        await b._connect_once("https://secure.example.com", "u1")

    assert captured["url"].startswith("wss://")
    assert isinstance(captured["kwargs"]["ssl"], ssl.SSLContext)
    assert captured["kwargs"]["ssl"].verify_mode == ssl.CERT_NONE


@pytest.mark.asyncio
async def test_wss_scheme_no_ssl_kwarg_when_disabled(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, Any] = {}

    def fake_connect(url: str, **kwargs: Any) -> _FakeConnect:
        captured["kwargs"] = kwargs
        return _FakeConnect(url, **kwargs)

    monkeypatch.setattr(bridgemod.websockets, "connect", fake_connect)

    b = bridgemod.ConnectorMcpBridge(
        server_name="s", tool_definitions=[], call_tool=lambda n, a: {},
        get_token=lambda: _async_none(), allow_private_certificate=False,
    )
    with pytest.raises(_StopTest):
        await b._connect_once("https://secure.example.com", "u1")

    assert "ssl" not in captured["kwargs"]


@pytest.mark.asyncio
async def test_connect_always_disables_compression(monkeypatch: pytest.MonkeyPatch) -> None:
    """2026-08-13 실제 XGEN dev 서버(dev-xgen.x2bee.com) 상대 실측 회귀 테스트.

    permessage-deflate(websockets 기본값)를 켠 채 접속하면 핸드셰이크는
    성공하지만 hello 전송 직후 서버 경로의 무언가(게이트웨이/프록시로
    추정)가 close 프레임 없이 연결을 끊었다. compression=None으로 끄니
    ready가 정상적으로 돌아왔다. 이 인자가 다시 빠지면 조용히 재발한다."""
    captured: dict[str, Any] = {}

    def fake_connect(url: str, **kwargs: Any) -> _FakeConnect:
        captured["kwargs"] = kwargs
        return _FakeConnect(url, **kwargs)

    monkeypatch.setattr(bridgemod.websockets, "connect", fake_connect)

    b = bridgemod.ConnectorMcpBridge(
        server_name="s", tool_definitions=[], call_tool=lambda n, a: {},
        get_token=lambda: _async_none(),
    )
    with pytest.raises(_StopTest):
        await b._connect_once("https://secure.example.com", "u1")

    assert captured["kwargs"]["compression"] is None


async def _async_none() -> None:
    return None
