"""ConnectorMcpBridge의 진짜 왕복을 검증한다 — 목(mock)으로 대충 흉내내는 게
아니라, aiohttp로 실제 로컬 TCP 소켓에 HTTP(로그인)와 WebSocket(도구 브릿지)
서버를 띄우고, xgen-seepage의 실제 클라이언트 코드(HttpClient/AuthApi/
ConnectorMcpBridge)가 그 서버와 실제로 통신하게 한다.

검증 흐름: login() → access_token 획득 → bridge가 WS로 붙어 hello 전송 →
서버가 ready로 응답 → 서버가 mcp_call(inspect_csv)을 보냄 → 브릿지가 진짜
xgen_seepage.tools.call_tool로 실행 → mcp_result로 정확한 결과를 돌려주는지
서버 쪽에서 확인.

이게 통과하면 "XGEN 서버를 바라봐서 에이전트 기능이 동작하게" 라는 요구사항의
와이어 레벨 왕복이 실제로 동작한다는 뜻이다(진짜 XGEN 백엔드가 문서화된
그대로 응답한다는 전제 하에 — 그 전제 자체는 실제 서버에 붙여서만 최종
검증 가능하다. ARCHITECTURE.md 참고)."""
from __future__ import annotations

import asyncio
import json
from pathlib import Path

import pytest
from aiohttp import web, WSMsgType

from xgen_seepage import csv_adapter
from xgen_seepage.connector.auth import AuthApi
from xgen_seepage.connector.bridge import ConnectorMcpBridge
from xgen_seepage.connector.http_client import HttpClient
from xgen_seepage.tools import TOOL_DEFINITIONS, call_tool

TEST_PASSWORD = "correct-password"
TEST_ACCESS_TOKEN = "tok-e2e"
TEST_USER_ID = "user-e2e"


def _sha256_hex(s: str) -> str:
    import hashlib
    return hashlib.sha256(s.encode()).hexdigest()


class MockXgenServer:
    """login + connector-mcp WS 엔드포인트만 흉내내는 최소 aiohttp 서버."""

    def __init__(self) -> None:
        self.received_call: dict | None = None
        self.ready_frame: dict | None = None
        self._server_authorization: str | None = None

    async def handle_login(self, request: web.Request) -> web.Response:
        body = await request.json()
        if body.get("password") == _sha256_hex(TEST_PASSWORD):
            return web.json_response({
                "success": True, "access_token": TEST_ACCESS_TOKEN,
                "refresh_token": "ref-e2e", "token_type": "bearer",
                "user_id": TEST_USER_ID, "username": body["email"],
            })
        return web.json_response({"success": False, "message": "bad creds", "access_token": None})

    async def handle_bridge_ws(self, request: web.Request) -> web.WebSocketResponse:
        self._server_authorization = request.headers.get("Authorization")
        ws = web.WebSocketResponse()
        await ws.prepare(request)

        # 1. hello 를 기다린다
        hello_raw = await ws.receive_str()
        hello = json.loads(hello_raw)
        assert hello["type"] == "hello"

        # 2. ready 로 ack
        await ws.send_str(json.dumps({
            "type": "ready", "catalog_id": hello["catalog_id"], "tool_count": len(hello["tools"]),
        }))
        self.ready_frame = hello

        # 3. 실제 도구 하나를 호출해본다 (inspect_csv)
        request_id = "req-1"
        csv_path = request.app["csv_path"]
        await ws.send_str(json.dumps({
            "type": "mcp_call", "request_id": request_id,
            "server": "xgen-seepage", "tool": "inspect_csv", "args": {"path": csv_path},
        }))

        # 4. mcp_result 를 받는다
        async for msg in ws:
            if msg.type == WSMsgType.TEXT:
                data = json.loads(msg.data)
                if data.get("type") == "mcp_result" and data.get("request_id") == request_id:
                    self.received_call = data
                    await ws.close()
                    break
        return ws


@pytest.fixture
async def mock_server(tmp_path: Path):
    csv_path = tmp_path / "e2e.csv"
    csv_path.write_text("name,dept\nkim,ai\n", encoding="utf-8")

    server_state = MockXgenServer()
    app = web.Application()
    app["csv_path"] = str(csv_path)
    app.router.add_post("/api/auth/login", server_state.handle_login)
    app.router.add_get("/api/tools/ws/connector-mcp/{user_id}", server_state.handle_bridge_ws)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "127.0.0.1", 0)
    await site.start()
    port = site._server.sockets[0].getsockname()[1]  # type: ignore[union-attr]

    yield f"http://127.0.0.1:{port}", server_state, csv_path
    await runner.cleanup()


@pytest.mark.asyncio
async def test_full_login_and_bridge_roundtrip(mock_server) -> None:
    base_url, server_state, csv_path = mock_server

    http = HttpClient(base_url)
    auth = AuthApi(http)
    login_result = await auth.login("tester@example.com", TEST_PASSWORD)
    assert login_result.access_token == TEST_ACCESS_TOKEN
    await http.aclose()

    async def get_token() -> str:
        return login_result.access_token

    bridge = ConnectorMcpBridge(
        server_name="xgen-seepage",
        tool_definitions=TOOL_DEFINITIONS,
        call_tool=call_tool,
        get_token=get_token,
    )

    run_task = asyncio.create_task(bridge.run(base_url, TEST_USER_ID))
    try:
        for _ in range(100):  # 최대 5초 폴링 — 서버가 mcp_result 를 받을 때까지
            if server_state.received_call is not None:
                break
            await asyncio.sleep(0.05)
    finally:
        bridge.stop()
        run_task.cancel()
        try:
            await run_task
        except asyncio.CancelledError:
            pass

    assert server_state.ready_frame is not None
    advertised = {t["name"] for t in server_state.ready_frame["tools"]}
    assert "inspect_csv" in advertised
    assert all(t["server"] == "xgen-seepage" for t in server_state.ready_frame["tools"])

    assert server_state.received_call is not None
    assert server_state.received_call["ok"] is True
    result = server_state.received_call["result"]
    assert result["name"] == "e2e.csv"
    assert result["preview"] == [["name", "dept"], ["kim", "ai"]]

    # 직접 파일도 같은 값이어야 한다 — 브릿지를 거치며 뭔가 변형되지 않았는지.
    direct = csv_adapter.get_schema(csv_path)
    assert direct.preview == result["preview"]
