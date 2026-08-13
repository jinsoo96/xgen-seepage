"""커넥터 MCP 브릿지.

XGEN 백엔드의 `/api/tools/ws/connector-mcp/{user_id}` WebSocket에 붙어 이
프로세스의 도구 카탈로그(xgen_seepage.tools.TOOL_DEFINITIONS)를 광고하고,
서버가 보내는 `mcp_call`을 `xgen_seepage.tools.call_tool`로 실행해
`mcp_result`로 돌려준다. 로그인한 사용자가 XGEN 어디서(웹 UI 등) 어떤
에이전트와 채팅하든, 그 에이전트가 이 프로세스에 붙은 xlwings/CSV 도구를
바로 호출할 수 있게 되는 지점이 여기다 — "(llm외) 에이전트들의 기능이
동작하게" 라는 요구사항을 채우는 핵심 컴포넌트.

와이어 프로토콜은 PlateerLab/xgen-connector `src/main/mcp-bridge.ts`가 쓰는
것과 **완전히 동일하다**(참고 전용 확인 — 이 프로젝트는 xgen-connector에
의존하지 않고 이 프로토콜을 Python으로 독립적으로 재구현한다. NOTICE 참조):

  클라이언트→서버: {"type":"hello","catalog_id":<str>,"tools":[...]}
                   {"type":"ping"}                              (20초 간격)
                   {"type":"mcp_result","request_id":..,"ok":..,"result"|"error":..}
  서버→클라이언트: {"type":"ready","catalog_id":..,"tool_count":..}
                   {"type":"mcp_call","request_id":..,"server":..,"tool":..,"args":..}
"""
from __future__ import annotations

import asyncio
import json
import logging
import time
from dataclasses import dataclass
from typing import Any, Awaitable, Callable
from urllib.parse import quote

import websockets
from websockets.exceptions import ConnectionClosed

log = logging.getLogger("xgen-seepage.bridge")

HEARTBEAT_SECONDS = 20
RECONNECT_MIN_SECONDS = 5
RECONNECT_MAX_SECONDS = 60

ToolCallHandler = Callable[[str, dict[str, Any]], dict[str, Any]]
TokenProvider = Callable[[], Awaitable[str | None]]


@dataclass
class BridgeStatus:
    connected: bool = False
    catalog_synced: bool = False
    server_tool_count: int = 0
    error: str | None = None


def ws_url(server_url: str, user_id: str) -> str:
    base = server_url.rstrip("/")
    if base.startswith("https://"):
        base = "wss://" + base[len("https://"):]
    elif base.startswith("http://"):
        base = "ws://" + base[len("http://"):]
    return f"{base}/api/tools/ws/connector-mcp/{quote(user_id, safe='')}"


class ConnectorMcpBridge:
    """`server_name`으로 광고되는 도구 하나를 XGEN 에이전트 세션에 연결한다."""

    def __init__(
        self,
        *,
        server_name: str,
        tool_definitions: list[dict[str, Any]],
        call_tool: ToolCallHandler,
        get_token: TokenProvider,
    ) -> None:
        self._server_name = server_name
        self._tool_definitions = tool_definitions
        self._call_tool = call_tool
        self._get_token = get_token
        self.status = BridgeStatus()
        self._stopped = True
        self._catalog_seq = 0
        self._pending_catalog_id = ""

    async def run(self, server_url: str, user_id: str) -> None:
        """접속을 유지하며 끊기면 지수 백오프로 재접속한다. stop()이
        호출될 때까지 반환하지 않는다 — 호출자가 태스크로 감싸 실행한다."""
        self._stopped = False
        backoff = RECONNECT_MIN_SECONDS
        while not self._stopped:
            try:
                await self._connect_once(server_url, user_id)
                backoff = RECONNECT_MIN_SECONDS
            except asyncio.CancelledError:
                raise
            except Exception as e:
                self.status = BridgeStatus(connected=False, error=str(e))
                log.warning("bridge disconnected: %s", e)
            if self._stopped:
                break
            await asyncio.sleep(backoff)
            backoff = min(RECONNECT_MAX_SECONDS, backoff * 1.8)

    def stop(self) -> None:
        self._stopped = True

    async def _connect_once(self, server_url: str, user_id: str) -> None:
        token = await self._get_token()
        headers = {"Authorization": f"Bearer {token}"} if token else {}
        url = ws_url(server_url, user_id)
        async with websockets.connect(url, additional_headers=headers) as ws:
            self.status = BridgeStatus(connected=True)
            await self._send_hello(ws)
            heartbeat = asyncio.create_task(self._heartbeat_loop(ws))
            try:
                async for raw in ws:
                    await self._on_message(ws, raw)
            finally:
                heartbeat.cancel()

    async def _heartbeat_loop(self, ws: Any) -> None:
        try:
            while True:
                await asyncio.sleep(HEARTBEAT_SECONDS)
                await ws.send(json.dumps({"type": "ping"}))
        except (asyncio.CancelledError, ConnectionClosed):
            return

    async def _send_hello(self, ws: Any) -> None:
        catalog_id = f"{int(time.time() * 1000)}-{self._catalog_seq}"
        self._catalog_seq += 1
        self._pending_catalog_id = catalog_id
        tools = [
            {
                "server": self._server_name,
                "name": t["name"],
                "description": t.get("description"),
                "inputSchema": t.get("input_schema"),
            }
            for t in self._tool_definitions
        ]
        await ws.send(json.dumps({"type": "hello", "catalog_id": catalog_id, "tools": tools}))

    async def _on_message(self, ws: Any, raw: str) -> None:
        try:
            msg = json.loads(raw)
        except json.JSONDecodeError:
            return
        kind = msg.get("type")
        if kind == "ready":
            if msg.get("catalog_id") != self._pending_catalog_id:
                return
            self.status = BridgeStatus(
                connected=True,
                catalog_synced=True,
                server_tool_count=int(msg.get("tool_count") or 0),
            )
            return
        if kind == "mcp_call":
            request_id = msg.get("request_id")
            tool = msg.get("tool")
            args = msg.get("args") or {}
            try:
                result = self._call_tool(str(tool), args)
                payload: dict[str, Any] = {"request_id": request_id, "ok": True, "result": result}
            except Exception as e:  # call_tool 자체는 예외를 안 던지지만 방어적으로
                payload = {"request_id": request_id, "ok": False, "error": str(e)}
            await ws.send(json.dumps({"type": "mcp_result", **payload}))
