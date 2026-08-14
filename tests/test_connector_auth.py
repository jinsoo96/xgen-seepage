"""auth.py를 httpx.MockTransport로 검증한다. 실제 네트워크 없이, XGEN
게이트웨이가 문서화된 대로 응답한다는 가정 하에 요청 바디/헤더/파싱 경로를
실제로 돈다(WebSocket 브릿지까지 도는 진짜 E2E는
tests/test_connector_bridge_integration.py)."""
from __future__ import annotations

import json

import httpx
import pytest

from xgen_seepage.connector.auth import AuthApi
from xgen_seepage.connector.hash import sha256_hex
from xgen_seepage.connector.http_client import ApiError, HttpClient

CORRECT_PASSWORD = "correct-horse-battery-staple"
CORRECT_HASH = sha256_hex(CORRECT_PASSWORD)


def _make_client(handler: httpx.MockTransport) -> HttpClient:
    async_client = httpx.AsyncClient(transport=handler, base_url="http://mock")
    return HttpClient("http://mock", client=async_client)


def _login_handler(request: httpx.Request) -> httpx.Response:
    body = json.loads(request.content)
    assert request.url.path == "/api/auth/login"
    # 반드시 해시로 보내야 한다. 평문을 그대로 보내면 실제 게이트웨이에서 항상 실패한다.
    if body["password"] == CORRECT_HASH and body["token"] is None:
        return httpx.Response(200, json={
            "success": True, "access_token": "tok-abc", "refresh_token": "ref-abc",
            "token_type": "bearer", "user_id": "u1", "username": body["email"],
        })
    return httpx.Response(200, json={"success": False, "message": "bad credentials", "access_token": None})


@pytest.mark.asyncio
async def test_login_sends_hashed_password_and_parses_tokens() -> None:
    transport = httpx.MockTransport(_login_handler)
    http = _make_client(transport)
    auth = AuthApi(http)

    result = await auth.login("me@example.com", CORRECT_PASSWORD)

    assert result.access_token == "tok-abc"
    assert result.refresh_token == "ref-abc"
    assert result.user_id == "u1"
    assert result.username == "me@example.com"
    await http.aclose()


@pytest.mark.asyncio
async def test_login_wrong_password_raises() -> None:
    transport = httpx.MockTransport(_login_handler)
    http = _make_client(transport)
    auth = AuthApi(http)

    with pytest.raises(RuntimeError, match="bad credentials"):
        await auth.login("me@example.com", "wrong-password")
    await http.aclose()


def _validate_handler(request: httpx.Request) -> httpx.Response:
    body = json.loads(request.content)
    if body["token"] == "tok-abc":
        return httpx.Response(200, json={
            "valid": True, "user_id": "u1", "username": "tester",
            "is_superuser": False, "roles": ["member"], "permissions": ["chat"],
        })
    return httpx.Response(200, json={"valid": False})


@pytest.mark.asyncio
async def test_validate_returns_current_user_for_valid_token() -> None:
    transport = httpx.MockTransport(_validate_handler)
    http = _make_client(transport)
    auth = AuthApi(http)

    user, new_access = await auth.validate("tok-abc")

    assert user is not None
    assert user.user_id == "u1"
    assert user.roles == ["member"]
    assert new_access is None
    await http.aclose()


@pytest.mark.asyncio
async def test_validate_returns_none_for_invalid_token() -> None:
    transport = httpx.MockTransport(_validate_handler)
    http = _make_client(transport)
    auth = AuthApi(http)

    user, _ = await auth.validate("expired-token")

    assert user is None
    await http.aclose()


@pytest.mark.asyncio
async def test_401_triggers_auth_failure_hook() -> None:
    def handler(_request: httpx.Request) -> httpx.Response:
        return httpx.Response(401, json={"detail": "unauthorized"})

    called = {"count": 0}

    async def on_auth_failure() -> None:
        called["count"] += 1

    transport = httpx.MockTransport(handler)
    async_client = httpx.AsyncClient(transport=transport, base_url="http://mock")
    http = HttpClient("http://mock", client=async_client, on_auth_failure=on_auth_failure)

    with pytest.raises(ApiError) as exc_info:
        await http.get("/api/agentflow/list")

    assert exc_info.value.status == 401
    assert called["count"] == 1
    await http.aclose()
