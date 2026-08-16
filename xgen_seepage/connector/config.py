"""xgen-seepage 자체 로컬 설정.

xgen-connector의 `connector.json`과는 **별개의 독립 파일**이다.
xgen-connector 설치 여부와 무관하게 동작해야 한다는 요구사항 때문에, 이
프로젝트만의 설정 파일/키체인 항목을 쓴다(파일이 두는 정보의 종류만
xgen-connector `src/main/config.ts`를 참고했다. NOTICE 참조).

토큰(access/refresh)은 이 JSON 파일에 평문으로 두지 않는다. OS 키체인
(Windows Credential Manager / macOS Keychain / Linux libsecret)에 `keyring`
패키지로 저장한다. xgen-connector가 JWT를 평문 파일에 두지 않고 OS
키체인에 두는 것과 같은 이유(설치 프로그램이 도난당해도 토큰이 파일
그대로 노출되지 않게 하기 위함)다.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import keyring

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
    # Excel 태스크팬 채팅이 매 메시지마다 실행하는 저장된 agentflow 워크플로우.
    # 비어 있으면 첫 `run` 시 자동 생성한다(connector/agentflow_client.py).
    chat_workflow_id: str = ""
    chat_provider: str = ""
    chat_model: str = ""
    taskpane_port: int = DEFAULT_TASKPANE_PORT

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
        chat_provider=data.get("chat_provider", ""),
        chat_model=data.get("chat_model", ""),
        taskpane_port=int(data.get("taskpane_port") or DEFAULT_TASKPANE_PORT),
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


def get_access_token() -> str | None:
    return _get_quietly(_KEY_ACCESS)


def get_refresh_token() -> str | None:
    return _get_quietly(_KEY_REFRESH)


def _get_quietly(key: str) -> str | None:
    try:
        return keyring.get_password(_SERVICE, key)
    except Exception:
        # 읽기 실패는 "토큰 없음"과 동일하게. 재로그인을 유도하면 된다.
        return None


def set_tokens(access_token: str | None, refresh_token: str | None) -> None:
    try:
        if access_token:
            keyring.set_password(_SERVICE, _KEY_ACCESS, access_token)
        else:
            _delete_quietly(_KEY_ACCESS)
        if refresh_token:
            keyring.set_password(_SERVICE, _KEY_REFRESH, refresh_token)
        else:
            _delete_quietly(_KEY_REFRESH)
    except Exception as e:
        raise KeyringUnavailableError(
            "OS 키체인에 토큰을 저장하지 못했습니다. 대화형 로그인 세션이 아닌 "
            "환경(서비스 계정·원격/자동화 실행 등)에서 실행 중이면 Windows "
            "Credential Manager/macOS Keychain에 접근할 수 없습니다. 일반 "
            f"사용자 데스크톱 세션에서 다시 실행하세요. (원인: {type(e).__name__}: {e})"
        ) from e


def clear_tokens() -> None:
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
