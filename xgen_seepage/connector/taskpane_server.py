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
from typing import Any, Awaitable, Callable

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


def _friendly_xgen_error(e: ApiError) -> JSONResponse:
    """XGEN이 준 에러 상태 코드를 패널에 실행 가능한 한국어 메시지로 바꿔
    준다. 특히 403은 토큰 문제가 아니라 **권한 문제**다(게이트웨이가 JWT로
    권한을 판정) - 재로그인해도 안 풀리니 서버/계정을 확인하라고 정확히
    알려준다. 그냥 "403"만 뜨면 사용자가 뭘
    해야 할지 알 수 없다."""
    if e.status == 403:
        message = (
            "이 계정은 이 서버에서 워크플로우(에이전트) 접근 권한이 없습니다"
            "(403 Forbidden). 재로그인해도 안 풀립니다 - 권한이 아니라 계정/서버 "
            "문제입니다. 권한 있는 계정인지, 지금 붙은 XGEN 서버(jeju/dev/prod)가 "
            "맞는지 확인하세요. 터미널에서 `xgen-seepage status`(권한 확인) / "
            "`xgen-seepage server list`(서버 확인·전환)."
        )
    elif e.status == 401:
        message = "로그인이 만료됐습니다(401). 터미널에서 `xgen-seepage login`으로 다시 로그인하세요."
    else:
        message = f"XGEN 서버 오류({e.status})."
    return JSONResponse(
        {"error": "xgen_error", "status": e.status, "message": message, "detail": e.body},
        status_code=e.status,
    )


class _ChatApiHandler:
    """`/workflows`(에이전트 목록)와 `/chat/stream`(선택한 에이전트 실행)
    라우트 본체. 클로저 대신 클래스로 둔 이유는 `TaskpaneServer`가 가진
    `get_token`/`agentflow`/`chat_workflow_id`를 매 요청마다 다시 읽어야
    해서다(토큰은 만료·갱신될 수 있고, config는 나중에 재로그인으로 바뀔 수
    있다). 바운드 메서드로 라우트에 등록해야 한다 - Starlette의 `Route`는
    `inspect.isfunction`/`ismethod`로 감지된 것만 `(request) -> Response`로
    취급하고, `__call__` 인스턴스는 raw ASGI 앱으로 오인한다(실측 버그)."""

    def __init__(
        self,
        *,
        get_token: TokenProvider,
        agentflow: AgentflowApi,
        get_chat_workflow_id: Callable[[], str],
        get_server_info: Callable[[], tuple[str, str]] | None = None,
    ) -> None:
        self._get_token = get_token
        self._agentflow = agentflow
        self._get_chat_workflow_id = get_chat_workflow_id
        self._get_server_info = get_server_info or (lambda: ("", ""))

    async def _authed(self) -> str | None:
        token = await self._get_token()
        if token is not None:
            self._agentflow.set_token(token)
        return token

    async def server_info(self, _request: Request):
        """패널이 지금 붙은 XGEN 서버와 로그인 계정을 보여주기 위한 정보.
        어느 서버에 붙었는지 안 보이면 "왜 403이지?"를 진단할 수 없다 -
        엉뚱한 서버(prod 대신 jeju 등)에 붙어 있는 걸 눈으로 잡게 한다."""
        server_url, username = self._get_server_info()
        return JSONResponse({"server_url": server_url, "username": username})

    async def list_workflows(self, _request: Request):
        """태스크팬 드롭다운이 채우는 에이전트(워크플로우) 목록. XGEN에
        로그인만 돼 있으면 그 계정이 가진 워크플로우 중 하나를 골라 바로
        붙일 수 있게 한다 - CLI로 미리 하나 박아둘 필요 없이 패널 안에서
        토글로 선택. 현재 config에 설정된 기본값도 함께 알려준다."""
        token = await self._authed()
        if token is None:
            return JSONResponse(
                {"error": "not_logged_in", "message": "`xgen-seepage login`이 필요합니다."},
                status_code=401,
            )
        try:
            workflows = await self._agentflow.list_workflows()
        except ApiError as e:
            return _friendly_xgen_error(e)
        return JSONResponse(
            {
                "workflows": [
                    {"workflow_id": w.workflow_id, "workflow_name": w.workflow_name}
                    for w in workflows
                ],
                "current": self._get_chat_workflow_id(),
            }
        )

    async def chat_stream(self, request: Request):
        token = await self._authed()
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
        # 패널에서 고른 에이전트를 우선 쓰고, 안 넘어오면 config 기본값으로 폴백.
        workflow_id = body.get("workflow_id") or self._get_chat_workflow_id()
        if not workflow_id:
            return JSONResponse(
                {
                    "error": "no_chat_workflow",
                    "message": "채팅할 에이전트를 선택하세요(위 드롭다운). "
                    "목록이 비어 있으면 XGEN 캔버스에서 워크플로우를 먼저 만드세요.",
                },
                status_code=503,
            )
        interaction_id = body.get("interaction_id") or "default"

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
            return _friendly_xgen_error(e)

        async def _combined():
            if first_chunk:
                yield first_chunk
            async for chunk in gen:
                yield chunk

        return StreamingResponse(_combined(), media_type="text/event-stream")


class _NoCacheStaticFiles(StaticFiles):
    """정적 응답에 `Cache-Control: no-cache`를 붙인다. 이게 없으면 브라우저·
    특히 Excel 태스크팬 웹뷰가 예전 index.html/taskpane.js를 캐시해, 패널을
    고쳐도 새로고침만으로는 반영이 안 된다. no-cache는 매번 서버에 재검증하게
    해(etag 그대로 활용) 바뀌었을 때만 새로 받는다."""

    async def get_response(self, path: str, scope: Any) -> Any:
        resp = await super().get_response(path, scope)
        resp.headers["Cache-Control"] = "no-cache"
        return resp


def _build_app(chat_handler: _ChatApiHandler | None) -> Starlette:
    routes: list[BaseRoute] = [
        Route("/health", _health),
    ]
    if chat_handler is not None:
        routes.append(Route("/server", chat_handler.server_info, methods=["GET"]))
        routes.append(Route("/workflows", chat_handler.list_workflows, methods=["GET"]))
        routes.append(Route("/chat/stream", chat_handler.chat_stream, methods=["POST"]))
    if _TASKPANE_DIR.is_dir():
        routes.append(Mount("/", app=_NoCacheStaticFiles(directory=str(_TASKPANE_DIR), html=True)))
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
        get_server_info: Callable[[], tuple[str, str]] | None = None,
    ) -> None:
        chat_handler = None
        self._chat_http: HttpClient | None = None
        # get_chat_workflow_id는 이제 선택. 안 넘겨도 패널에서 에이전트를
        # 고를 수 있으니 빈 기본값으로 둔다(로그인 토큰과 서버 URL만 있으면
        # /workflows·/chat/stream을 붙인다).
        if xgen_server_url and get_token is not None:
            self._chat_http = HttpClient(xgen_server_url, allow_private_certificate=allow_private_certificate)
            chat_handler = _ChatApiHandler(
                get_token=get_token,
                agentflow=AgentflowApi(self._chat_http),
                get_chat_workflow_id=get_chat_workflow_id or (lambda: ""),
                get_server_info=get_server_info,
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
