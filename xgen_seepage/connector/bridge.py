"""커넥터 MCP 브릿지.

XGEN 백엔드의 `/api/tools/ws/connector-mcp/{user_id}` WebSocket에 붙어 이
프로세스의 도구 카탈로그(xgen_seepage.tools.TOOL_DEFINITIONS)를 광고하고,
서버가 보내는 `mcp_call`을 `xgen_seepage.tools.call_tool`로 실행해
`mcp_result`로 돌려준다. 로그인한 사용자가 XGEN 어디서(웹 UI 등) 어떤
에이전트와 채팅하든, 그 에이전트가 이 프로세스에 붙은 xlwings/CSV 도구를
바로 호출할 수 있게 되는 지점이 여기다. "(llm외) 에이전트들의 기능이
동작하게" 라는 요구사항을 채우는 핵심 컴포넌트.

서버와 주고받는 와이어 프로토콜(hello/ready/ping/mcp_call/mcp_result)은
다음과 같다:

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

from .connection_security import default_ssl_context, relaxed_ssl_context

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
        allow_private_certificate: bool = False,
    ) -> None:
        self._server_name = server_name
        self._tool_definitions = tool_definitions
        self._call_tool = call_tool
        self._get_token = get_token
        self._allow_private_certificate = allow_private_certificate
        self.status = BridgeStatus()
        self._stopped = True
        self._catalog_seq = 0
        self._pending_catalog_id = ""
        self._current_ws: Any = None

    async def run(self, server_url: str, user_id: str) -> None:
        """접속을 유지하며 끊기면 지수 백오프로 재접속한다. stop()이
        호출될 때까지 반환하지 않는다. 호출자가 태스크로 감싸 실행한다."""
        self._stopped = False
        backoff: float = RECONNECT_MIN_SECONDS
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
        """실측(2026-08-17, 실제 Excel 머신)으로 찾은 버그: `_stopped` 플래그만
        세우면 재접속 사이 대기 구간에서만 체크된다 - `_connect_once`의
        `async for raw in ws:` 수신 루프가 메시지를 기다리며 블록돼 있는
        동안엔 아무도 안 깨워서 `run()`이 영원히 안 끝난다(정상 종료 경로가
        사실상 hang). 지금 연결된 소켓이 있으면 직접 닫아서 그 루프를
        깨운다 - `close()`는 코루틴이라 동기 메서드 안에서 스케줄만 건다
        (호출 시점에 이벤트 루프가 돌고 있다는 전제, 이 클래스의 모든 실사용
        경로가 그렇다)."""
        self._stopped = True
        if self._current_ws is not None:
            asyncio.ensure_future(self._current_ws.close())

    async def _connect_once(self, server_url: str, user_id: str) -> None:
        token = await self._get_token()
        headers = {"Authorization": f"Bearer {token}"} if token else {}
        url = ws_url(server_url, user_id)
        # TLS 컨텍스트는 wss:// 에만 의미가 있다(ws:// 에 넘기면 websockets가
        # 바로 에러). 폐쇄망 내부 CA를 위해 기본도 OS 신뢰 저장소 컨텍스트를
        # 명시적으로 넘긴다(HttpClient와 동일 정책, connection_security 참조).
        # allow_private_certificate면 검증 전면 해제.
        ssl_kwarg: dict[str, Any] = {}
        if url.startswith("wss://"):
            ssl_kwarg["ssl"] = (
                relaxed_ssl_context()
                if self._allow_private_certificate
                else default_ssl_context()
            )
        # compression=None: 실측(2026-08-13, 실제 XGEN dev 서버)에서 기본값인
        # permessage-deflate 확장을 켠 채로 접속하면 핸드셰이크는 성공하지만
        # hello를 보낸 직후 서버 쪽 프록시/게이트웨이가 close 프레임 없이
        # 연결을 끊었다. 압축을 끄니 즉시 정상적으로 ready가 돌아왔다 -
        # 이 경로의 중간 인프라가 압축 확장을 못 받는 것으로 보인다.
        async with websockets.connect(
            url, additional_headers=headers, compression=None, **ssl_kwarg
        ) as ws:
            self.status = BridgeStatus(connected=True)
            self._current_ws = ws
            await self._send_hello(ws)
            heartbeat = asyncio.create_task(self._heartbeat_loop(ws))
            try:
                async for raw in ws:
                    await self._on_message(ws, raw)
            finally:
                heartbeat.cancel()
                self._current_ws = None

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

    async def _on_message(self, ws: Any, raw: str | bytes) -> None:
        try:
            msg = json.loads(raw)  # json.loads는 str/bytes 둘 다 받는다
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
            log.info(
                "도구 카탈로그 동기화 완료: 보낸 %d개, 서버 등록 %d개",
                len(self._tool_definitions), self.status.server_tool_count,
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
