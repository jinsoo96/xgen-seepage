"""xgen-seepage 자체 로컬 설정.

이 프로젝트만의 독립 설정 파일이며, 별도 데스크톱 클라이언트의 설치
여부와 무관하게 동작한다.

토큰(access/refresh)은 가능하면 이 JSON 파일에 두지 않고 OS 키체인
(Windows Credential Manager / macOS Keychain / Linux libsecret)에 `keyring`
패키지로 저장한다. JWT를 평문 파일에 두지 않고 OS 키체인에 두는 이유는
설치 프로그램이 도난당해도 토큰이 파일 그대로 노출되지 않게 하기
위함이다.

단, 헤드리스/서비스 세션·원격(SSH) 실행·폐쇄망 서버 계정처럼 **OS 키체인
자체에 접근할 수 없는 환경**이 실제로 존재한다(macOS는 GUI 로그인 세션이
아니면 -25308, Windows는 로그온 세션이 없으면 WinError 1312). 이런 곳에서
키체인을 강제하면 "어디서든 파이썬 의존성 걱정 없이 돌아가야 한다"는 이
프로젝트의 핵심 요구가 깨진다. 그래서 키체인이 불가할 때만 설정 폴더의
`tokens.json`(권한 0600, 소유자 전용)으로 **폴백**한다 - 키체인이 되는
데스크톱에서는 여전히 키체인이 우선이고, 파일 폴백은 키체인이 없는
환경에서의 최후 수단이다(트레이드오프: 파일 노출 시 토큰 노출).
"""
from __future__ import annotations

import json
import logging
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import keyring
import keyring.errors

_log = logging.getLogger(__name__)

_SERVICE = "xgen-seepage"
_KEY_ACCESS = "access_token"
_KEY_REFRESH = "refresh_token"


def config_dir() -> Path:
    return Path.home() / ".xgen-seepage"


def _config_path() -> Path:
    return config_dir() / "config.json"


# 태스크팬 로컬 서버 기본 포트. manifest.xml의 SourceLocation에 고정 값으로
# 박아 넣으므로(Phase 4) 런마다 바뀌면 안 된다 - 바꾸려면 매니페스트
# 재배포가 필요한 breaking change로 취급한다.
DEFAULT_TASKPANE_PORT = 51837


@dataclass
class SeepageConfig:
    server_url: str = ""
    user_id: str = ""
    username: str = ""
    allow_private_certificate: bool = False
    # Excel 태스크팬 채팅이 매 메시지마다 실행하는 XGEN 워크플로우의 id.
    # xgen-seepage가 새로 만들지 않는다 - 사용자가 XGEN에 이미 갖고 있는
    # 워크플로우(agents/harness 노드 포함) 중 하나에 연결하는 것.
    # `xgen-seepage chat-workflow list`/`set`로 관리한다.
    chat_workflow_id: str = ""
    taskpane_port: int = DEFAULT_TASKPANE_PORT
    # 지금까지 로그인해 본 XGEN 서버 URL들. 다음 `login` 때 목록으로 보여줘
    # 다시 타이핑하지 않고 고를 수 있게 한다(jeju/dev/prod 등 여러 XGEN).
    # 고객사 내부 주소를 레포에 하드코딩하지 않기 위해, 프리셋이 아니라
    # "실제로 써 본 서버"만 로컬 config에 쌓는다.
    known_servers: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def load_config() -> SeepageConfig:
    p = _config_path()
    if not p.exists():
        return SeepageConfig()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return SeepageConfig()
    return SeepageConfig(
        server_url=data.get("server_url", ""),
        user_id=data.get("user_id", ""),
        username=data.get("username", ""),
        allow_private_certificate=bool(data.get("allow_private_certificate", False)),
        chat_workflow_id=data.get("chat_workflow_id", ""),
        taskpane_port=int(data.get("taskpane_port") or DEFAULT_TASKPANE_PORT),
        known_servers=list(data.get("known_servers") or []),
    )


def save_config(config: SeepageConfig) -> None:
    config_dir().mkdir(parents=True, exist_ok=True)
    _config_path().write_text(
        json.dumps(config.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
    )


class KeyringUnavailableError(RuntimeError):
    """OS 키체인(Windows Credential Manager 등)에 접근할 수 없다.

    실제로 관측된 원인(2026-08-13, PyInstaller로 얼린 exe를 대화형 데스크톱
    로그온 세션이 아닌 컨텍스트에서 실행했을 때): `keyring`의 Windows 백엔드가
    내부적으로 CredRead를 먼저 호출하는데, 세션이 없으면 win32 `WinError 1312
    ("지정된 로그온 세션이 없습니다")`를 raw로 던진다. keyring.errors 계층으로
    감싸지 않은 채 그대로 새어나온다. 서비스 계정/원격 자동화 실행 등에서 같은
    문제가 재현될 수 있어 여기서 붙잡아 실행 가능한 메시지로 바꾼다.
    """


# 토큰은 서버 URL별로 따로 저장한다. 여러 XGEN(jeju/dev/prod)을 오갈 때
# 한 슬롯을 덮어쓰면, 활성 서버는 B인데 키체인엔 A의 토큰이 남아 B에 A
# 토큰을 보내 403이 나는 사고가 생긴다(실측). 서버별로 키를 나눠 그 불일치
# 자체가 생기지 않게 한다. 키 형식: "access_token::<server_url>".
def _norm_server(server_url: str) -> str:
    return (server_url or "").strip().rstrip("/")


def _access_key(server_url: str) -> str:
    return f"{_KEY_ACCESS}::{_norm_server(server_url)}"


def _refresh_key(server_url: str) -> str:
    return f"{_KEY_REFRESH}::{_norm_server(server_url)}"


def get_access_token(server_url: str) -> str | None:
    v = _get_quietly(_access_key(server_url))
    if v is None:
        # 예전 버전은 서버 구분 없이 단일 슬롯에 저장했다. 그 토큰을 현재
        # 서버 것으로 간주해 읽어 준다(다음 login/set_tokens에서 서버별
        # 슬롯으로 옮겨가고 레거시 슬롯은 지운다).
        return _get_quietly(_KEY_ACCESS)
    return v


def get_refresh_token(server_url: str) -> str | None:
    v = _get_quietly(_refresh_key(server_url))
    if v is None:
        return _get_quietly(_KEY_REFRESH)
    return v


# ── 파일 폴백(키체인 불가 환경 전용) ─────────────────────────────────────
# 키체인에 못 쓰는 헤드리스/폐쇄망에서만 쓰는 소유자 전용(0600) JSON.
def _tokens_file() -> Path:
    return config_dir() / "tokens.json"


def _file_read(key: str) -> str | None:
    p = _tokens_file()
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8")).get(key)
    except (json.JSONDecodeError, OSError):
        return None


def _file_write(key: str, value: str | None) -> None:
    p = _tokens_file()
    try:
        data = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    except (json.JSONDecodeError, OSError):
        data = {}
    if value:
        data[key] = value
    else:
        data.pop(key, None)
    if not data:
        try:
            p.unlink()
        except OSError:
            pass
        return
    config_dir().mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    try:
        os.chmod(p, 0o600)  # 소유자만 읽기/쓰기
    except OSError:
        pass


def _get_quietly(key: str) -> str | None:
    try:
        v = keyring.get_password(_SERVICE, key)
    except Exception:
        # 읽기 실패는 "키체인 없음"과 동일하게 보고 파일 폴백을 확인한다.
        v = None
    if v is not None:
        return v
    return _file_read(key)


def _keyring_write(key: str, value: str | None) -> None:
    """keyring에 한 항목을 쓰거나 지운다. 백엔드 자체가 불가하면 예외를
    그대로 위로 던져(set_tokens가 파일 폴백으로 전환하게) 한다."""
    if value:
        keyring.set_password(_SERVICE, key, value)
    else:
        try:
            keyring.delete_password(_SERVICE, key)
        except keyring.errors.PasswordDeleteError:
            pass  # 원래 없던 항목


def set_tokens(server_url: str, access_token: str | None, refresh_token: str | None) -> None:
    ak, rk = _access_key(server_url), _refresh_key(server_url)
    try:
        _keyring_write(ak, access_token)
        _keyring_write(rk, refresh_token)
        # 서버별 슬롯에 확실히 옮겨 담았으니 레거시 단일 슬롯은 정리한다.
        _delete_quietly(_KEY_ACCESS)
        _delete_quietly(_KEY_REFRESH)
        # 키체인 저장 성공 → 예전에 남았을 수 있는 파일 폴백 잔재를 지운다
        # (두 곳에 동시에 남아 나중에 stale 토큰을 읽는 사고 방지).
        _file_write(ak, None)
        _file_write(rk, None)
    except Exception as e:
        # 키체인 자체가 불가한 환경(헤드리스/서비스/원격/폐쇄망) → 파일(0600) 폴백.
        _log.warning(
            "OS 키체인에 접근할 수 없어 토큰을 %s (0600)에 저장합니다. "
            "키체인이 되는 데스크톱 세션에서는 키체인이 우선입니다. (원인: %s: %s)",
            _tokens_file(), type(e).__name__, e,
        )
        _file_write(ak, access_token)
        _file_write(rk, refresh_token)
        _file_write(_KEY_ACCESS, None)
        _file_write(_KEY_REFRESH, None)


def has_token(server_url: str) -> bool:
    """이 서버에 저장된 access token이 있는가(레거시 단일 슬롯 폴백 없이,
    이 서버 슬롯만 본다). 서버 목록에서 '토큰 있음/없음'을 정확히 보여줄 때
    쓴다 - get_access_token은 마이그레이션용 레거시 폴백이 있어 여러 서버가
    다 '있음'으로 보일 수 있다."""
    return _get_quietly(_access_key(server_url)) is not None


def clear_tokens(server_url: str) -> None:
    _delete_quietly(_access_key(server_url))
    _delete_quietly(_refresh_key(server_url))
    # 레거시 단일 슬롯도 같이 정리(마이그레이션 잔재 방지).
    _delete_quietly(_KEY_ACCESS)
    _delete_quietly(_KEY_REFRESH)


def _delete_quietly(key: str) -> None:
    try:
        keyring.delete_password(_SERVICE, key)
    except Exception:
        # PasswordDeleteError(원래 없어서 못 지움)뿐 아니라 백엔드 자체가 깨진
        # 경우(위 KeyringUnavailableError와 같은 원인)도 "이미 없는 셈" 취급한다.
        # 정리/로그아웃 경로가 키체인 상태 때문에 실패해서는 안 된다.
        pass
    # 파일 폴백에 남은 항목도 함께 정리(로그아웃/클리어가 한쪽만 지우지 않게).
    _file_write(key, None)
