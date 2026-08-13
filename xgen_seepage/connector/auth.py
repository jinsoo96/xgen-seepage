"""XGEN 게이트웨이 인증.

흐름(login → {access_token, refresh_token}, 이후 모든 호출은 Authorization:
Bearer <access_token>, validate-token으로 신원 확인, refresh로 재발급)은
PlateerLab/xgen-connector `src/core/auth.ts`와 같은 API 계약을 따르되
(참고 전용 확인 — 실제 엔드포인트/필드명은 XGEN 백엔드가 정본이고 이 파일이
그걸 재발명하지 않는다는 뜻이다. NOTICE 참조), Python으로 새로 구현했다.
"""
from __future__ import annotations

from dataclasses import dataclass

from .hash import sha256_hex
from .http_client import HttpClient


@dataclass
class LoginResult:
    access_token: str
    refresh_token: str | None
    token_type: str
    user_id: str
    username: str


@dataclass
class CurrentUser:
    user_id: str
    username: str
    is_superuser: bool
    roles: list[str]
    permissions: list[str]


class AuthApi:
    def __init__(self, http: HttpClient) -> None:
        self._http = http

    async def login(self, email: str, password: str) -> LoginResult:
        """이메일 + 평문 비밀번호로 로그인한다. 비밀번호는 전송 전에
        SHA-256 hex로 해싱한다(게이트웨이가 해시값을 그대로 비교함)."""
        password_hash = sha256_hex(password)
        res = await self._http.post(
            "/api/auth/login",
            {"email": email, "password": password_hash, "token": None},
            auth=False,
        )
        if not res.get("success") or not res.get("access_token"):
            raise RuntimeError(res.get("message") or "로그인에 실패했습니다.")
        return LoginResult(
            access_token=res["access_token"],
            refresh_token=res.get("refresh_token"),
            token_type=res.get("token_type", "bearer"),
            user_id=res.get("user_id", ""),
            username=res.get("username", email),
        )

    async def validate(
        self, access_token: str, refresh_token: str | None = None
    ) -> tuple[CurrentUser | None, str | None]:
        """토큰을 검증하고 현재 사용자를 반환한다. 만료됐고 refresh_token이
        있으면 (None, new_access_token) 형태로 재발급 토큰을 받을 수 있다."""
        res = await self._http.post(
            "/api/auth/validate-token",
            {"token": access_token, "refresh_token": refresh_token},
            auth=False,
        )
        if not res.get("valid"):
            return None, res.get("new_access_token")
        user = CurrentUser(
            user_id=res.get("user_id", ""),
            username=res.get("username", ""),
            is_superuser=bool(res.get("is_superuser")),
            roles=res.get("roles") or [],
            permissions=res.get("permissions") or [],
        )
        return user, res.get("new_access_token")

    async def refresh(self, refresh_token: str) -> str | None:
        res = await self._http.post(
            "/api/auth/refresh", {"refresh_token": refresh_token}, auth=False
        )
        return res.get("access_token") if res.get("success") else None

    async def logout(self, access_token: str) -> None:
        try:
            await self._http.post("/api/auth/logout", {"token": access_token}, timeout=8.0)
        except Exception:
            pass  # 로그아웃은 best-effort
