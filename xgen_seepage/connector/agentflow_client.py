"""XGEN agentflow API - 태스크팬 채팅이 매 메시지마다 실행하는 저장된
워크플로우를 만들고 돌리는 데 쓴다.

**하네스 전용 CRUD를 쓴다, 레거시 `/api/agentflow/save`가 아니다**: 실측
(2026-08-17)으로 확인한 실제 이유가 있다. `agents/harness` 노드가 든
워크플로우를 레거시 `/api/agentflow/save`로 저장하면 `{"success": true}`가
돌아오지만, 그 직후 `/api/agentflow/load/{id}`·`/api/agentflow/list`·
`/api/agentflow/harness/workflows` 어디로도 안 보였다(즉 어디 갔는지 알 수
없는 상태가 됨). 백엔드 소스(`xgen-workflow controller/workflow/endpoints/
harness.py:1183-1191`)의 주석이 정확히 이 부류의 문제를 설명한다: 레거시
저장 경로는 `(user_id, workflow_id, workflow_name)` 3컬럼 존재 확인이라
같은 workflow_id에 이름만 갈리면 insert 분기로 새 줄이 쌓이는 버그가 있고,
레거시 delete는 `LIMIT 1`이라 중복 줄을 온전히 못 지운다. 그래서 하네스
워크플로우 전용으로 새로 만든 `/api/agentflow/harness/workflows` CRUD가
"duplicate-proof"로 명시돼 있다 - 이쪽만 쓴다.
"""
from __future__ import annotations

import json as jsonlib
from dataclasses import dataclass
from typing import Any, AsyncIterator

from .http_client import ApiError, HttpClient

# 채팅 스트림은 에이전트가 생각/도구호출/생성하는 동안 오래 걸릴 수 있다.
# 일반 JSON 호출(로그인 등)의 30초 기본 타임아웃과는 별개로 넉넉히 둔다.
_STREAM_TIMEOUT_SECONDS = 300.0


@dataclass
class HarnessWorkflow:
    workflow_id: str
    workflow_name: str
    raw: dict[str, Any]


class AgentflowApi:
    def __init__(self, http: HttpClient) -> None:
        self._http = http

    def set_token(self, token: str | None) -> None:
        self._http.set_token(token)

    async def load_harness_workflow(self, workflow_id: str) -> HarnessWorkflow | None:
        """존재하면 반환, 없으면(404) None. 그 외 에러는 그대로 올린다."""
        try:
            res = await self._http.get(f"/api/agentflow/harness/workflows/{workflow_id}")
        except ApiError as e:
            if e.status == 404:
                return None
            raise
        return HarnessWorkflow(
            workflow_id=res.get("workflow_id", workflow_id),
            workflow_name=res.get("workflow_name", workflow_id),
            raw=res,
        )

    async def save_harness_workflow(
        self, workflow_id: str, workflow_name: str, workflow_data: dict[str, Any]
    ) -> HarnessWorkflow:
        res = await self._http.post(
            "/api/agentflow/harness/workflows",
            {
                "workflow_id": workflow_id,
                "workflow_name": workflow_name,
                "workflow_data": workflow_data,
            },
        )
        return HarnessWorkflow(workflow_id=workflow_id, workflow_name=workflow_name, raw=res)

    async def execute_stream(
        self,
        *,
        workflow_id: str,
        workflow_name: str,
        input_data: str | None = None,
        interaction_id: str = "default",
    ) -> AsyncIterator[bytes]:
        """`POST /api/agentflow/execute/based-id/stream`의 원시 SSE 바이트를
        그대로 흘려보낸다(파싱/재직렬화 없음 - `taskpane_server.py`가 이걸
        그대로 클라이언트에게 다시 흘려보내는 프록시라서, XGEN이 실제로
        주는 이벤트 형태를 이 계층에서 왜곡하면 안 된다)."""
        body = {
            "workflow_id": workflow_id,
            "workflow_name": workflow_name,
            "input_data": input_data,
            "interaction_id": interaction_id,
            "response_format": "stream",
        }
        async with self._http.stream(
            "POST", "/api/agentflow/execute/based-id/stream", body, timeout=_STREAM_TIMEOUT_SECONDS
        ) as resp:
            if resp.status_code >= 400:
                raw = await resp.aread()
                try:
                    detail = jsonlib.loads(raw).get("detail", raw.decode("utf-8", "replace"))
                except (jsonlib.JSONDecodeError, UnicodeDecodeError):
                    detail = raw.decode("utf-8", "replace")
                raise ApiError(resp.status_code, f"execute_stream -> {resp.status_code}", detail)
            async for chunk in resp.aiter_raw():
                yield chunk
