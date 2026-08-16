"""Excel 태스크팬용 로컬 HTTPS 서버.

`xgen-seepage run`이 이미 상주시키는 `ConnectorMcpBridge`(XGEN 서버로 나가는
WS 연결)와 나란히, 반대 방향(로컬에서 들어오는 HTTPS)으로 하나 더 띄운다.
태스크팬(Office.js, Excel 안에 도킹된 패널)이 이 서버에 붙는다.

책임은 딱 두 가지다:
  1. 태스크팬 정적 자산(HTML/JS/CSS/manifest) 서빙 - 같은 origin이라 CORS가
     아예 생기지 않는다.
  2. `/chat/stream` - 이 커넥터가 이미 갖고 있는 토큰으로 실제 XGEN의
     `POST /api/agentflow/execute/based-id/stream`을 대신 호출하고, 그
     SSE 원시 바이트를 그대로 흘려보낸다(재파싱/재직렬화 없음 - XGEN이
     실제로 보내는 이벤트 형태를 왜곡할 이유가 없다).

셀 IO는 이 서버가 하지 않는다 - 그건 여전히 전적으로 `bridge.py`가 이미
하는 일이다(에이전트의 tool call → WS → `tools.call_tool()` →
`live_adapter`/`libreoffice_adapter`). 태스크팬은 그 결과를 화면으로
보여주는 채팅창일 뿐, 셀을 직접 만지지 않는다.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Awaitable, Callable

import uvicorn
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, StreamingResponse
from starlette.routing import BaseRoute, Mount, Route
from starlette.staticfiles import StaticFiles

from .agentflow_client import AgentflowApi
from .http_client import ApiError, HttpClient

log = logging.getLogger("xgen-seepage.taskpane")

_TASKPANE_DIR = Path(__file__).resolve().parent.parent / "taskpane"

TokenProvider = Callable[[], Awaitable[str | None]]


async def _health(_request: Request) -> JSONResponse:
    return JSONResponse({"ok": True, "service": "xgen-seepage-taskpane"})


class _ChatStreamHandler:
    """`/chat/stream` 라우트 본체. 클로저 대신 클래스로 둔 이유는 `TaskpaneServer`가
    가진 `get_token`/`agentflow`/`chat_workflow_id`를 매 요청마다 다시 읽어야
    해서다(토큰은 만료·갱신될 수 있고, config는 나중에 재로그인으로 바뀔 수 있다)."""

    def __init__(
        self,
        *,
        get_token: TokenProvider,
        agentflow: AgentflowApi,
        get_chat_workflow_id: Callable[[], str],
    ) -> None:
        self._get_token = get_token
        self._agentflow = agentflow
        self._get_chat_workflow_id = get_chat_workflow_id

    async def handle(self, request: Request):
        """일반 함수가 아니라 바운드 메서드로 등록해야 한다 - Starlette의
        `Route`는 `inspect.isfunction`/`ismethod`로 감지된 것만 `(request) ->
        Response` 시그니처로 취급해서 감싼다. 임의의 콜러블 인스턴스(예:
        이 클래스에 `__call__`을 뒀다면)는 raw ASGI 앱(`scope, receive,
        send`)으로 오인해서 호출한다(실측으로 확인한 진짜 버그: `__call__`로
        뒀더니 `takes 2 positional arguments but 4 were given`)."""
        workflow_id = self._get_chat_workflow_id()
        if not workflow_id:
            return JSONResponse(
                {"error": "no_chat_workflow", "message": "채팅 워크플로우가 아직 설정되지 않았습니다."},
                status_code=503,
            )
        token = await self._get_token()
        if token is None:
            return JSONResponse(
                {"error": "not_logged_in", "message": "`xgen-seepage login`이 필요합니다."},
                status_code=401,
            )
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "bad_request", "message": "JSON 본문이 필요합니다."}, status_code=400)
        message = body.get("message")
        if not message:
            return JSONResponse({"error": "bad_request", "message": "'message' 필드가 필요합니다."}, status_code=400)
        interaction_id = body.get("interaction_id") or "default"

        self._agentflow.set_token(token)
        gen = self._agentflow.execute_stream(
            workflow_id=workflow_id,
            workflow_name=workflow_id,
            input_data=message,
            interaction_id=interaction_id,
        )
        # StreamingResponse는 첫 send에서 상태 코드를 확정한다. XGEN 쪽 에러
        # (예: 워크플로우 없음)는 스트림을 열자마자 알 수 있으므로, 첫 청크를
        # 미리 하나 당겨서 에러면 제대로 된 상태 코드로 응답하고, 아니면
        # 그 청크를 포함해 나머지를 정상적으로 흘려보낸다.
        try:
            first_chunk = await gen.__anext__()
        except StopAsyncIteration:
            first_chunk = b""
        except ApiError as e:
            return JSONResponse(
                {"error": "xgen_error", "status": e.status, "detail": e.body}, status_code=e.status
            )

        async def _combined():
            if first_chunk:
                yield first_chunk
            async for chunk in gen:
                yield chunk

        return StreamingResponse(_combined(), media_type="text/event-stream")


def _build_app(chat_handler: _ChatStreamHandler | None) -> Starlette:
    routes: list[BaseRoute] = [
        Route("/health", _health),
    ]
    if chat_handler is not None:
        routes.append(Route("/chat/stream", chat_handler.handle, methods=["POST"]))
    if _TASKPANE_DIR.is_dir():
        routes.append(Mount("/", app=StaticFiles(directory=str(_TASKPANE_DIR), html=True)))
    else:
        log.warning("taskpane 정적 자산 폴더가 없습니다: %s (정적 서빙 비활성)", _TASKPANE_DIR)
    return Starlette(routes=routes)


class TaskpaneServer:
    """`bridge.py`의 `ConnectorMcpBridge`와 같은 run()/stop() 모양을 따른다 -
    `cmd_run`이 둘을 `asyncio.gather`로 나란히 돌릴 수 있게."""

    def __init__(
        self,
        *,
        host: str = "127.0.0.1",
        port: int,
        cert_path: Path,
        key_path: Path,
        xgen_server_url: str | None = None,
        allow_private_certificate: bool = False,
        get_token: TokenProvider | None = None,
        get_chat_workflow_id: Callable[[], str] | None = None,
    ) -> None:
        chat_handler = None
        self._chat_http: HttpClient | None = None
        if xgen_server_url and get_token is not None and get_chat_workflow_id is not None:
            self._chat_http = HttpClient(xgen_server_url, allow_private_certificate=allow_private_certificate)
            chat_handler = _ChatStreamHandler(
                get_token=get_token,
                agentflow=AgentflowApi(self._chat_http),
                get_chat_workflow_id=get_chat_workflow_id,
            )

        app = _build_app(chat_handler)
        config = uvicorn.Config(
            app,
            host=host,
            port=port,
            ssl_certfile=str(cert_path),
            ssl_keyfile=str(key_path),
            log_level="warning",
        )
        self._server = uvicorn.Server(config)
        self.host = host
        self.port = port

    async def run(self) -> None:
        """`stop()`이 불릴 때까지 반환하지 않는다. 호출자가 태스크로 감싸 실행한다."""
        await self._server.serve()

    def stop(self) -> None:
        self._server.should_exit = True

    async def aclose(self) -> None:
        if self._chat_http is not None:
            await self._chat_http.aclose()
