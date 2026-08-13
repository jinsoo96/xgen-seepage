"""XGEN 게이트웨이용 최소 비동기 HTTP 클라이언트.

책임 분담(베이스 URL 정규화, Bearer 헤더 자동 부착, JSON 요청/응답, 401에서
auth-failure 훅)은 PlateerLab/xgen-connector `src/core/client.ts`(HttpClient)
와 같은 모양을 따르되(참고 전용, NOTICE 참조), httpx 기반으로 새로 구현했다.
"""
from __future__ import annotations

from typing import Any, Awaitable, Callable

import httpx


class ApiError(Exception):
    def __init__(self, status: int, message: str, body: Any = None) -> None:
        super().__init__(message)
        self.status = status
        self.body = body


def normalize_base_url(url: str) -> str:
    return (url or "").strip().rstrip("/")


class HttpClient:
    def __init__(
        self,
        base_url: str,
        *,
        on_auth_failure: Callable[[], Awaitable[None] | None] | None = None,
        timeout: float = 30.0,
        client: httpx.AsyncClient | None = None,
    ) -> None:
        self._base_url = normalize_base_url(base_url)
        self._token: str | None = None
        self._on_auth_failure = on_auth_failure
        self._timeout = timeout
        self._client = client or httpx.AsyncClient(timeout=timeout)

    def set_base_url(self, base_url: str) -> None:
        self._base_url = normalize_base_url(base_url)

    def get_base_url(self) -> str:
        return self._base_url

    def set_token(self, token: str | None) -> None:
        self._token = token

    def _url(self, path: str) -> str:
        return f"{self._base_url}{path if path.startswith('/') else '/' + path}"

    def _headers(self, extra: dict[str, str] | None = None) -> dict[str, str]:
        headers = dict(extra or {})
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        return headers

    async def _handle_auth_failure(self) -> None:
        if self._on_auth_failure is None:
            return
        result = self._on_auth_failure()
        if result is not None:
            await result

    async def json(
        self,
        method: str,
        path: str,
        body: Any = None,
        *,
        auth: bool = True,
        timeout: float | None = None,
    ) -> Any:
        headers = self._headers({"Content-Type": "application/json", "Accept": "application/json"})
        resp = await self._client.request(
            method,
            self._url(path),
            headers=headers,
            json=body if body is not None else None,
            timeout=timeout or self._timeout,
        )
        parsed: Any = None
        if resp.content:
            try:
                parsed = resp.json()
            except ValueError:
                parsed = resp.text
        if resp.status_code >= 400:
            if resp.status_code == 401 and auth:
                await self._handle_auth_failure()
            raise ApiError(resp.status_code, f"{method} {path} -> {resp.status_code}", parsed)
        return parsed

    async def get(self, path: str, **kw: Any) -> Any:
        return await self.json("GET", path, **kw)

    async def post(self, path: str, body: Any = None, **kw: Any) -> Any:
        return await self.json("POST", path, body, **kw)

    async def aclose(self) -> None:
        await self._client.aclose()
