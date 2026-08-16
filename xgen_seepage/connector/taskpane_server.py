"""Excel 태스크팬용 로컬 HTTPS 서버.

`xgen-seepage run`이 이미 상주시키는 `ConnectorMcpBridge`(XGEN 서버로 나가는
WS 연결)와 나란히, 반대 방향(로컬에서 들어오는 HTTPS)으로 하나 더 띄운다.
태스크팬(Office.js, Excel 안에 도킹된 패널)이 이 서버에 붙는다.

책임은 딱 두 가지로 좁힌다:
  1. 태스크팬 정적 자산(HTML/JS/CSS/manifest) 서빙 - 같은 origin이라 CORS가
     아예 생기지 않는다.
  2. (Phase 2에서 추가) `/chat/stream` - 이 커넥터가 이미 갖고 있는 토큰으로
     실제 XGEN의 에이전트 실행 스트림을 프록시한다.

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
from starlette.responses import JSONResponse
from starlette.routing import BaseRoute, Mount, Route
from starlette.staticfiles import StaticFiles

log = logging.getLogger("xgen-seepage.taskpane")

_TASKPANE_DIR = Path(__file__).resolve().parent.parent / "taskpane"

TokenProvider = Callable[[], Awaitable[str | None]]


async def _health(_request):
    return JSONResponse({"ok": True, "service": "xgen-seepage-taskpane"})


def _build_app() -> Starlette:
    routes: list[BaseRoute] = [
        Route("/health", _health),
    ]
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
    ) -> None:
        app = _build_app()
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
