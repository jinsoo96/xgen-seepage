"""xgen-seepage 커넥터 CLI.

설치 후 최초 1회 `xgen-seepage login`으로 XGEN 서버에 로그인해 토큰을 OS
키체인에 저장하고, `xgen-seepage run`으로 백그라운드 브릿지를 띄운다.
그 이후로는 사용자가 XGEN 어디서(웹 UI 등) 어떤 에이전트와 채팅하든, 그
에이전트가 이 프로세스에 붙은 xlwings/CSV 도구를 자동으로 쓸 수 있다.

비대화형 배포(폐쇄망 대량 설치)를 위해 `login`은 인자/환경변수로도 값을
받는다: `--server`/`XGEN_SEEPAGE_SERVER_URL`, `--email`/`XGEN_SEEPAGE_EMAIL`,
`--password`/`XGEN_SEEPAGE_PASSWORD`(권장하지 않음. 가능하면 대화형 입력).
"""
from __future__ import annotations

import argparse
import asyncio
import getpass
import logging
import os
import sys
import uuid

from . import certs, config as cfgmod
from .auth import AuthApi
from .bridge import ConnectorMcpBridge
from .http_client import ApiError, HttpClient
from .taskpane_server import TaskpaneServer

log = logging.getLogger("xgen-seepage")

# 도구 카탈로그를 광고할 때 쓰는 고정 서버 이름. hello 프레임의 "server" 필드.
SERVER_NAME = "xgen-seepage"


def _env(name: str) -> str | None:
    return os.environ.get(f"XGEN_SEEPAGE_{name}")


def _env_bool(name: str) -> bool | None:
    """"true"/"1"/"false"/"0" 파싱. xgen-connector deployment-defaults.ts의
    optionalBoolean과 같은 규칙(참고 전용, NOTICE 참조)."""
    raw = _env(name)
    if raw is None or raw == "":
        return None
    if raw.lower() in ("true", "1"):
        return True
    if raw.lower() in ("false", "0"):
        return False
    raise ValueError(f"XGEN_SEEPAGE_{name} must be true/false/1/0, got {raw!r}")


async def cmd_login(args: argparse.Namespace) -> int:
    server_url = args.server or _env("SERVER_URL")
    email = args.email or _env("EMAIL")
    password = args.password or _env("PASSWORD")
    allow_private_certificate = args.allow_private_certificate
    if allow_private_certificate is None:
        allow_private_certificate = _env_bool("ALLOW_PRIVATE_CERTIFICATE") or False

    if not server_url:
        server_url = input("XGEN 서버 URL (예: https://xgen.example.com): ").strip()
    if not email:
        email = input("이메일: ").strip()
    if not password:
        password = getpass.getpass("비밀번호: ")

    http = HttpClient(server_url, allow_private_certificate=allow_private_certificate)
    auth = AuthApi(http)
    try:
        result = await auth.login(email, password)
    except (ApiError, RuntimeError) as e:
        print(f"로그인 실패: {e}", file=sys.stderr)
        return 1
    finally:
        await http.aclose()

    cfg = cfgmod.SeepageConfig(
        server_url=server_url,
        user_id=result.user_id,
        username=result.username,
        allow_private_certificate=allow_private_certificate,
    )
    cfgmod.save_config(cfg)
    try:
        cfgmod.set_tokens(result.access_token, result.refresh_token)
    except cfgmod.KeyringUnavailableError as e:
        print(f"로그인 자체는 성공했지만 토큰 저장에 실패했습니다: {e}", file=sys.stderr)
        return 1
    print(f"로그인 성공: {result.username} ({server_url})")
    return 0


async def cmd_logout(_args: argparse.Namespace) -> int:
    cfgmod.clear_tokens()
    print("로그아웃했습니다(토큰 삭제). 설정 파일(server_url 등)은 유지됩니다.")
    return 0


async def cmd_status(_args: argparse.Namespace) -> int:
    cfg = cfgmod.load_config()
    if not cfg.server_url:
        print("설정된 서버가 없습니다. `xgen-seepage login`을 먼저 실행하세요.")
        return 1
    access = cfgmod.get_access_token()
    print(f"서버: {cfg.server_url}")
    print(f"사용자: {cfg.username or '(미상)'} ({cfg.user_id or '?'})")
    if not access:
        print("토큰 없음. `xgen-seepage login` 필요")
        return 1
    http = HttpClient(cfg.server_url, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)
    try:
        user, _ = await auth.validate(access, cfgmod.get_refresh_token())
    except ApiError as e:
        print(f"검증 실패: {e}")
        return 1
    finally:
        await http.aclose()
    if user is None:
        print("토큰 만료/무효. `xgen-seepage login` 다시 필요")
        return 1
    print(f"토큰 유효. 역할: {user.roles}, 권한 {len(user.permissions)}개")
    return 0


async def _ensure_valid_token(cfg: cfgmod.SeepageConfig, auth: AuthApi) -> str | None:
    """유효한 access_token을 반환한다. 만료됐으면 refresh_token으로 재발급
    시도 후 키체인을 갱신한다. 둘 다 실패하면 None(재로그인 필요)."""
    access = cfgmod.get_access_token()
    if not access:
        return None
    refresh_token = cfgmod.get_refresh_token()
    try:
        user, new_access = await auth.validate(access, refresh_token)
    except ApiError:
        return None
    if user is not None:
        return access
    if new_access:
        cfgmod.set_tokens(new_access, refresh_token)
        return new_access
    if refresh_token:
        try:
            refreshed = await auth.refresh(refresh_token)
        except ApiError:
            refreshed = None
        if refreshed:
            cfgmod.set_tokens(refreshed, refresh_token)
            return refreshed
    return None


async def cmd_run(args: argparse.Namespace) -> int:
    cfg = cfgmod.load_config()
    if not cfg.server_url:
        print("설정된 서버가 없습니다. `xgen-seepage login`을 먼저 실행하세요.", file=sys.stderr)
        return 1

    http = HttpClient(cfg.server_url, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)

    async def get_token() -> str | None:
        return await _ensure_valid_token(cfg, auth)

    from .. import tools  # 순환 import 회피용 지연 import

    user_id = cfg.user_id or str(uuid.uuid4())
    bridge = ConnectorMcpBridge(
        server_name=SERVER_NAME,
        tool_definitions=tools.TOOL_DEFINITIONS,
        call_tool=tools.call_tool,
        get_token=get_token,
        allow_private_certificate=cfg.allow_private_certificate,
    )

    token = await get_token()
    if token is None:
        print("토큰이 없거나 만료됐습니다. `xgen-seepage login`을 다시 실행하세요.", file=sys.stderr)
        await http.aclose()
        return 1

    tasks = [asyncio.create_task(bridge.run(cfg.server_url, user_id))]
    taskpane: TaskpaneServer | None = None
    if not args.no_taskpane:
        cert_path, key_path = certs.ensure_dev_certificate(cfgmod.config_dir())
        taskpane = TaskpaneServer(
            port=cfg.taskpane_port, cert_path=cert_path, key_path=key_path
        )
        tasks.append(asyncio.create_task(taskpane.run()))
        print(f"태스크팬 로컬 서버: https://127.0.0.1:{taskpane.port}")

    print(f"{cfg.server_url} 에 연결합니다 (사용자 {cfg.username or user_id}). Ctrl+C로 종료.")
    try:
        await asyncio.gather(*tasks)
    except KeyboardInterrupt:
        pass
    finally:
        bridge.stop()
        if taskpane is not None:
            taskpane.stop()
        await asyncio.gather(*tasks, return_exceptions=True)
        await http.aclose()
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="xgen-seepage", description="XGEN Excel/CSV 커넥터")
    sub = p.add_subparsers(dest="command", required=True)

    login_p = sub.add_parser("login", help="XGEN 서버에 로그인하고 토큰을 저장한다")
    login_p.add_argument("--server", help="XGEN 서버 URL")
    login_p.add_argument("--email", help="로그인 이메일")
    login_p.add_argument("--password", help="비밀번호(권장하지 않음. 대화형 입력 사용)")
    login_p.add_argument(
        "--allow-private-certificate",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="사설/자체서명 CA를 쓰는 서버 허용(폐쇄망 내부 인증서). 기본값은 거부",
    )
    login_p.set_defaults(func=cmd_login)

    logout_p = sub.add_parser("logout", help="저장된 토큰을 삭제한다")
    logout_p.set_defaults(func=cmd_logout)

    status_p = sub.add_parser("status", help="현재 설정/토큰 상태를 확인한다")
    status_p.set_defaults(func=cmd_status)

    run_p = sub.add_parser("run", help="에이전트 도구 브릿지를 시작한다(포그라운드, 상주)")
    run_p.add_argument(
        "--no-taskpane",
        action="store_true",
        default=False,
        help="Excel 태스크팬용 로컬 HTTPS 서버를 띄우지 않는다(기본: 띄움)",
    )
    run_p.set_defaults(func=cmd_run)

    return p


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = build_parser()
    args = parser.parse_args()
    return asyncio.run(args.func(args))


if __name__ == "__main__":
    sys.exit(main())
