"""XGEN agentflow API - Excel 태스크팬 채팅이 쓴다.

**설계를 바로잡음(2026-08-17, 사용자 정정)**: 처음엔 "채팅용 워크플로우가
없으면 xgen-seepage가 하나 자동 생성한다"는 방향으로 만들고 있었다. 틀린
방향이었다 - 이미 XGEN에 있는 워크플로우에 **연결**하는 것이어야 한다.
사용자가 XGEN에 이미 갖고 있는 워크플로우(캔버스로 직접 만든 것,
`agents/harness` 노드가 든 것) 중 아무거나 하나에 태스크팬을 연결하는
것이지, xgen-seepage가
전용 워크플로우를 서버에 몰래 만들어서 소유하는 게 아니다. 그래서 이
모듈은 조회/실행만 하고, "이 워크플로우가 없으면 만든다" 같은 로직이
없다. 사용자가 `xgen-seepage chat-workflow list`/`set`으로 직접 고른다.

이 방향이 실측으로도 맞았다: 새로 만든 워크플로우는(하네스 전용 API로
만들든 레거시로 만들든) 만든 직후 조회/삭제 어디서도 안 보이는 실제
플랫폼 버그를 발견했다(아래 `list_workflows` 문서 참조) - 반면 사용자가
이미 갖고 있는 실제 워크플로우(`shinhan_blue_agent_v1`)에 연결해서 돌리니
그 문제를 완전히 비켜가고 실행이 끝까지 정상적으로 흘렀다(vLLM 모델
미기동이라는, 이미 §8에서 알려진 별개의 이슈만 남았다).

`save_harness_workflow`/`load_harness_workflow`는 "레거시 `/api/agentflow/
save`가 아니라 하네스 전용 CRUD를 쓴다"는 이유로 남겨뒀다 - 레거시 저장
경로는 `(user_id, workflow_id, workflow_name)` 3컬럼 존재 확인이라 같은
workflow_id에 이름만 갈리면 insert 분기로 중복 줄이 쌓이는 버그가 있고
(백엔드 소스 `xgen-workflow controller/workflow/endpoints/harness.py:
1183-1191`의 주석이 이 문제를 정확히 설명한다), 레거시 delete는
`LIMIT 1`이라 그 중복을 온전히 못 지운다.
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


@dataclass
class WorkflowSummary:
    workflow_id: str
    workflow_name: str


@dataclass
class ProviderOption:
    provider: str
    default_model: str


class AgentflowApi:
    def __init__(self, http: HttpClient) -> None:
        self._http = http
        # workflow_id -> [harness node id]. 워크플로우 구조는 자주 안 바뀌므로
        # provider 주입 때마다 다시 로드하지 않게 캐시한다.
        self._harness_nodes_cache: dict[str, list[str]] = {}

    def set_token(self, token: str | None) -> None:
        self._http.set_token(token)

    async def list_providers(self) -> list[ProviderOption]:
        """서버에 설정된 LLM provider와 그 기본 모델. 실측(2026-08-17)으로
        확인한 실제 문제: 워크플로우의 harness provider가 비어 있으면 서버
        기본값으로 떨어지는데 그 모델(vLLM 등)이 안 떠 있으면 실행이
        `503 No model loaded`로 죽는다. 그래서 패널에서 provider/model을 직접
        골라 실행 시 harness 노드에 주입할 수 있게, 여기서 후보 목록을 준다.
        `/api/config/persistent`의 카테고리별 `*_MODEL_DEFAULT`/`*_MODEL_NAME`
        을 읽어 구성한다."""
        res = await self._http.get("/api/config/persistent")
        cats = res.get("categories", {}) if isinstance(res, dict) else {}
        out: list[ProviderOption] = []
        # (provider 이름, 그 카테고리에서 기본 모델을 담은 설정 키 후보들)
        specs = [
            ("anthropic", ["ANTHROPIC_MODEL_DEFAULT"]),
            ("openai", ["OPENAI_MODEL_DEFAULT", "OPENAI_MODEL_NAME"]),
            ("gemini", ["GEMINI_MODEL_DEFAULT", "GEMINI_MODEL_NAME"]),
            ("vllm", ["VLLM_MODEL_NAME", "VLLM_MODEL_DEFAULT"]),
            ("deepseek", ["DEEPSEEK_MODEL_DEFAULT"]),
            ("bedrock", ["BEDROCK_MODEL_DEFAULT"]),
        ]
        for provider, model_keys in specs:
            configs = cats.get(provider, {}).get("configs", {})
            if not configs:
                continue
            model = ""
            for mk in model_keys:
                v = configs.get(mk, {}).get("current_value")
                if v:
                    model = str(v)
                    break
            out.append(ProviderOption(provider=provider, default_model=model))
        return out

    async def _harness_node_ids(self, workflow_id: str) -> list[str]:
        """워크플로우 안 `agents/harness` 노드들의 id. provider/model을 실행
        시점에 이 노드들로 주입하기 위해 필요하다(node_parameter는 node_id
        키다). 로드 실패하면 빈 리스트(주입 없이 서버 기본값으로 진행)."""
        if workflow_id in self._harness_nodes_cache:
            return self._harness_nodes_cache[workflow_id]
        ids: list[str] = []
        for path in (
            f"/api/agentflow/harness/workflows/{workflow_id}",
            f"/api/agentflow/load/{workflow_id}",
        ):
            try:
                res = await self._http.get(path)
            except ApiError:
                continue
            wd = res.get("workflow_data") or res.get("content") or res
            nodes = wd.get("nodes") if isinstance(wd, dict) else None
            if not nodes:
                continue
            for n in nodes:
                ntype = (n.get("data", {}) or {}).get("id") or n.get("type") or ""
                if ntype == "agents/harness":
                    nid = n.get("id")
                    if nid:
                        ids.append(nid)
            if ids:
                break
        self._harness_nodes_cache[workflow_id] = ids
        return ids

    async def list_workflows(self) -> list[WorkflowSummary]:
        """사용자가 태스크팬 채팅을 연결할 수 있는 워크플로우 후보 목록.
        `GET /api/agentflow/list`(일반 저장 워크플로우) 기준이다 - 실측으로
        확인한 실제 이유: `/api/agentflow/harness/workflows`로 방금 막
        만든 워크플로우는 같은 계정에서 곧바로 조회해도 안 보이는 문제가
        재현됐지만(원인: 서버 쪽, xgen-seepage 밖의 문제), 이미 캔버스로
        만들어 쓰고 있던 실제 워크플로우는 이 일반 목록에 정상적으로
        보이고 실행도 끝까지 됐다. 그래서 "새로 만들기"가 아니라 "이미
        있는 것 중 고르기"가 기본 경로다."""
        res = await self._http.get("/api/agentflow/list")
        raw_list = res.get("workflows", []) if isinstance(res, dict) else res
        return [
            WorkflowSummary(
                workflow_id=w.get("workflow_id", ""), workflow_name=w.get("workflow_name", "")
            )
            for w in raw_list
        ]

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
        provider: str | None = None,
        model: str | None = None,
    ) -> AsyncIterator[bytes]:
        """`POST /api/agentflow/execute/based-id/stream`의 원시 SSE 바이트를
        그대로 흘려보낸다(파싱/재직렬화 없음 - `taskpane_server.py`가 이걸
        그대로 클라이언트에게 다시 흘려보내는 프록시라서, XGEN이 실제로
        주는 이벤트 형태를 이 계층에서 왜곡하면 안 된다).

        `provider`/`model`을 주면 워크플로우의 모든 `agents/harness` 노드에
        `node_parameter`로 주입한다. 실측(2026-08-17)으로 이게 `503 No model
        loaded`의 해법임을 확인했다: provider가 비면 서버 기본값(안 떠 있는
        vLLM)으로 떨어져 죽는데, `provider=anthropic`을 주입하니 에이전트가
        실제로 도구를 호출하고 열린 Excel의 셀을 바꿨다."""
        # 바디 필드는 XGEN 실행 요청 형식을 따른다. include_tool_events를 켜야
        # 에이전트가 도구를 호출할 때 패널에 "[도구 호출] set_live_cell" 같은
        # 이벤트가 뜬다(사용자가 셀 편집이 실제로 일어나는 걸 눈으로 본다).
        body: dict[str, Any] = {
            "workflow_id": workflow_id,
            "workflow_name": workflow_name,
            "input_data": input_data,
            "interaction_id": interaction_id,
            "selected_collections": [],
            "selected_files": [],
            "include_logs": True,
            "include_node_status": True,
            "include_tool_events": True,
            "response_format": "stream",
        }
        if provider:
            node_ids = await self._harness_node_ids(workflow_id)
            override: dict[str, Any] = {"provider": provider}
            if model:
                override["model"] = model
            if node_ids:
                body["node_parameter"] = {nid: dict(override) for nid in node_ids}
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
