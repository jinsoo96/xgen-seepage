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
import platform
import subprocess
import sys
import uuid
import webbrowser
from pathlib import Path

from . import certs, config as cfgmod
from .agentflow_client import AgentflowApi
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


def _pick_server(known: list[str]) -> str:
    """이미 로그인해 본 XGEN 서버가 있으면 목록으로 보여주고 번호로 고르게
    한다(여러 XGEN - jeju/dev/prod 등을 오갈 때 다시 타이핑 안 하도록).
    없거나 '새 서버'를 고르면 직접 입력받는다."""
    if known:
        print("로그인할 XGEN 서버를 고르세요:")
        for i, url in enumerate(known, 1):
            print(f"  {i}. {url}")
        print(f"  {len(known) + 1}. 새 서버 직접 입력")
        choice = input("번호: ").strip()
        if choice.isdigit():
            n = int(choice)
            if 1 <= n <= len(known):
                return known[n - 1]
    return input("XGEN 서버 URL (예: https://xgen.example.com): ").strip()


async def cmd_login(args: argparse.Namespace) -> int:
    server_url = args.server or _env("SERVER_URL")
    email = args.email or _env("EMAIL")
    password = args.password or _env("PASSWORD")
    allow_private_certificate = args.allow_private_certificate
    if allow_private_certificate is None:
        allow_private_certificate = _env_bool("ALLOW_PRIVATE_CERTIFICATE") or False

    existing = cfgmod.load_config()
    if not server_url:
        server_url = _pick_server(existing.known_servers)
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

    # 토큰을 먼저 저장한다(서버별 슬롯). 저장에 실패하면 config의 활성
    # 서버를 바꾸지 않고 끝낸다 - 그래야 "활성 서버는 B인데 토큰은 없음/A의
    # 것"인 어긋난 상태가 남지 않는다.
    try:
        cfgmod.set_tokens(server_url, result.access_token, result.refresh_token)
    except cfgmod.KeyringUnavailableError as e:
        print(f"로그인 자체는 성공했지만 토큰 저장에 실패했습니다: {e}", file=sys.stderr)
        return 1
    # 방금 로그인한 서버를 known_servers 맨 앞에 넣어 다음 login 때 목록에 뜨게.
    known = [server_url] + [s for s in existing.known_servers if s != server_url]
    cfg = cfgmod.SeepageConfig(
        server_url=server_url,
        user_id=result.user_id,
        username=result.username,
        allow_private_certificate=allow_private_certificate,
        chat_workflow_id=existing.chat_workflow_id,
        taskpane_port=existing.taskpane_port,
        known_servers=known,
    )
    cfgmod.save_config(cfg)
    print(f"로그인 성공: {result.username} ({server_url})")
    return 0


async def cmd_logout(_args: argparse.Namespace) -> int:
    cfg = cfgmod.load_config()
    cfgmod.clear_tokens(cfg.server_url)
    print("로그아웃했습니다(토큰 삭제). 설정 파일(server_url 등)은 유지됩니다.")
    return 0


async def cmd_status(_args: argparse.Namespace) -> int:
    cfg = cfgmod.load_config()
    if not cfg.server_url:
        print("설정된 서버가 없습니다. `xgen-seepage login`을 먼저 실행하세요.")
        return 1
    access = cfgmod.get_access_token(cfg.server_url)
    print(f"서버: {cfg.server_url}")
    print(f"사용자: {cfg.username or '(미상)'} ({cfg.user_id or '?'})")
    if not access:
        print("토큰 없음. `xgen-seepage login` 필요")
        return 1
    http = HttpClient(cfg.server_url, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)
    try:
        user, _ = await auth.validate(access, cfgmod.get_refresh_token(cfg.server_url))
    except ApiError as e:
        print(f"검증 실패: {e}")
        return 1
    finally:
        await http.aclose()
    if user is None:
        print("토큰 만료/무효. `xgen-seepage login` 다시 필요")
        return 1
    print(f"토큰 유효. 슈퍼유저: {user.is_superuser}, 역할: {user.roles}, 권한 {len(user.permissions)}개")
    # 403의 실제 원인을 미리 알려준다: 워크플로우(에이전트) 조회/실행은
    # 게이트웨이가 이 계정의 권한으로 판정하며, 권한이 없으면 401이 아니라
    # 403 Forbidden으로 막힌다. 슈퍼유저가 아니고 agentflow 권한도 없으면
    # 그 서버에선 패널이 에이전트 목록을 못 불러오거나 실행이 막힌다.
    if not user.is_superuser and not _has_agentflow_access(user.permissions):
        print(
            "  주의: 이 계정은 이 서버에서 워크플로우(에이전트) 접근 권한이 없어 보입니다"
            "(agentflow read 권한 없음). 패널에서 403 Forbidden이 날 수 있습니다 - "
            "권한 있는 계정으로 로그인하거나(다른 서버면 `xgen-seepage server use`), "
            "관리자에게 권한을 요청하세요."
        )
    if len(cfg.known_servers) > 1:
        print(f"로그인해 본 서버 {len(cfg.known_servers)}개: `xgen-seepage server list`로 확인/전환")
    return 0


def _has_agentflow_access(permissions: list[str]) -> bool:
    """워크플로우(에이전트) 조회/실행에 필요한 권한을 갖고 있는가. XGEN
    게이트웨이가 요구하는 권한 이름 기준(참고: xgen-connector PROTOCOL.md의
    `main.agentflow:read`). 권한 표기가 서버 버전마다 조금씩 달라 넉넉히 매칭한다."""
    return any("agentflow" in p for p in permissions)


async def cmd_server_list(_args: argparse.Namespace) -> int:
    """지금까지 로그인해 본 XGEN 서버(jeju/dev/prod 등)와 각 서버의 토큰
    보유 여부를 보여준다. `server use`로 재타이핑 없이 전환한다."""
    cfg = cfgmod.load_config()
    servers = list(cfg.known_servers) or ([cfg.server_url] if cfg.server_url else [])
    if not servers:
        print("로그인해 본 서버가 없습니다. `xgen-seepage login`부터 하세요.")
        return 0
    for i, url in enumerate(servers, 1):
        marks = []
        if url == cfg.server_url:
            marks.append("현재")
        marks.append("토큰 있음" if cfgmod.has_token(url) else "토큰 없음")
        print(f"  {i}. {url}  [{', '.join(marks)}]")
    print("\n전환: `xgen-seepage server use <번호|URL>` (토큰 없으면 login 필요)")
    print("`run`이 떠 있으면 서버를 바꾼 뒤 재시작해야 반영됩니다.")
    return 0


async def cmd_server_use(args: argparse.Namespace) -> int:
    """활성 XGEN 서버를 바꾼다. 그 서버에 유효한 토큰이 이미 있으면 비밀번호
    없이 바로 전환하고, 없으면 `login`을 안내한다. xgen-connector처럼 서버
    전환은 계정 공간 전환이라, 토큰은 서버별로 분리해 둔다(한 서버 토큰이
    다른 서버로 새어 403이 나지 않게)."""
    cfg = cfgmod.load_config()
    servers = list(cfg.known_servers) or ([cfg.server_url] if cfg.server_url else [])
    target = (args.server or "").strip()
    if target.isdigit() and servers:
        n = int(target)
        if 1 <= n <= len(servers):
            target = servers[n - 1]
    target = target.rstrip("/")
    if not target:
        print("전환할 서버 URL 또는 번호를 지정하세요. `xgen-seepage server list`로 확인.", file=sys.stderr)
        return 1

    known = [target] + [s for s in servers if s != target]
    if not cfgmod.has_token(target):
        cfg.server_url = target
        cfg.known_servers = known
        cfgmod.save_config(cfg)
        print(f"활성 서버를 {target}로 바꿨습니다. 이 서버엔 저장된 토큰이 없으니 `xgen-seepage login` 하세요.")
        return 0

    # 저장된 토큰이 있으면 검증해서 이 서버 기준의 user_id/username으로 갱신한다
    # (서버마다 사용자 식별자가 다르므로, 브릿지 ws url에 옛 서버 user_id를
    # 쓰면 안 된다).
    cfg.server_url = target
    cfg.known_servers = known
    http = HttpClient(target, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)
    try:
        token = await _ensure_valid_token(cfg, auth)
        if token is None:
            cfgmod.save_config(cfg)
            print(f"활성 서버를 {target}로 바꿨지만 저장된 토큰이 만료됐습니다. `xgen-seepage login` 하세요.")
            return 0
        user, _ = await auth.validate(token, cfgmod.get_refresh_token(target))
        if user is not None:
            cfg.user_id = user.user_id or cfg.user_id
            cfg.username = user.username or cfg.username
        cfgmod.save_config(cfg)
    finally:
        await http.aclose()
    print(f"활성 서버를 {target}로 바꿨습니다(저장된 토큰 유효, 사용자 {cfg.username}).")
    print("`run`이 떠 있으면 재시작해야 반영됩니다.")
    return 0


async def cmd_chat_workflow_list(_args: argparse.Namespace) -> int:
    """Excel 태스크팬 채팅을 연결할 수 있는, 이미 XGEN에 있는 워크플로우
    목록을 보여준다. xgen-seepage가 새로 만들어주지 않는다 - 캔버스에서
    이미 만든 워크플로우(agents/harness 노드가 든 것) 중에서 고른다."""
    cfg = cfgmod.load_config()
    if not cfg.server_url:
        print("설정된 서버가 없습니다. `xgen-seepage login`을 먼저 실행하세요.", file=sys.stderr)
        return 1
    http = HttpClient(cfg.server_url, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)
    token = await _ensure_valid_token(cfg, auth)
    if token is None:
        print("토큰이 없거나 만료됐습니다. `xgen-seepage login`을 다시 실행하세요.", file=sys.stderr)
        await http.aclose()
        return 1
    http.set_token(token)
    agentflow = AgentflowApi(http)
    try:
        workflows = await agentflow.list_workflows()
    except ApiError as e:
        print(f"목록 조회 실패: {e}", file=sys.stderr)
        return 1
    finally:
        await http.aclose()
    if not workflows:
        print("워크플로우가 없습니다. XGEN 캔버스에서 먼저 하나 만드세요(agents/harness 노드 포함).")
        return 0
    current = cfg.chat_workflow_id
    for w in workflows:
        marker = " (현재 연결됨)" if w.workflow_id == current else ""
        print(f"  {w.workflow_id}  -  {w.workflow_name}{marker}")
    print("\n`xgen-seepage chat-workflow set <workflow_id>`로 태스크팬 채팅을 연결하세요.")
    return 0


async def cmd_chat_workflow_set(args: argparse.Namespace) -> int:
    cfg = cfgmod.load_config()
    if not cfg.server_url:
        print("설정된 서버가 없습니다. `xgen-seepage login`을 먼저 실행하세요.", file=sys.stderr)
        return 1
    http = HttpClient(cfg.server_url, allow_private_certificate=cfg.allow_private_certificate)
    auth = AuthApi(http)
    token = await _ensure_valid_token(cfg, auth)
    if token is None:
        print("토큰이 없거나 만료됐습니다. `xgen-seepage login`을 다시 실행하세요.", file=sys.stderr)
        await http.aclose()
        return 1
    http.set_token(token)
    agentflow = AgentflowApi(http)
    try:
        workflows = await agentflow.list_workflows()
    finally:
        await http.aclose()
    if not any(w.workflow_id == args.workflow_id for w in workflows):
        print(
            f"'{args.workflow_id}'가 워크플로우 목록에 없습니다. "
            "`xgen-seepage chat-workflow list`로 정확한 id를 확인하세요.",
            file=sys.stderr,
        )
        return 1
    cfg.chat_workflow_id = args.workflow_id
    cfgmod.save_config(cfg)
    print(f"태스크팬 채팅을 '{args.workflow_id}'에 연결했습니다.")
    return 0


_WEF_DEVELOPER_KEY = r"HKCU\Software\Microsoft\Office\16.0\WEF\Developer"


async def cmd_install_excel_addin(args: argparse.Namespace) -> int:
    """Excel 리본에 XGEN 버튼(태스크팬 채팅)을 띄운다. **폐쇄망에서 동작한다** -
    Microsoft 클라우드(스토어/애드인 삽입 서비스)를 전혀 안 거치는, 완전히
    로컬인 개발자 레지스트리 사이드로드다. 실측(2026-08-17, 실제 Office 2021
    볼륨 라이선스)으로 이 방식이 리본에 버튼을 실제로 띄우는 걸 확인했다
    (클라우드 애드인 삽입 대화상자는 폐쇄망/구형 빌드에서 콘텐츠를 못 그려
    못 쓴다 - 이건 그 경로를 안 탄다).

    하는 일: (1) 자체서명 인증서 생성 + 신뢰 등록(WebView2가 로컬 HTTPS
    페이지를 경고 없이 로드하도록), (2) 매니페스트를 실제 포트로 패치해 로컬에
    설치, (3) `HKCU\\...\\WEF\\Developer`에 매니페스트 경로를 등록. 값 이름도
    매니페스트 경로여야 한다(실측: 이름을 다른 문자열로 넣으면 Office가 안
    읽는다)."""
    if platform.system() != "Windows":
        print("Excel 리본 애드인 설치는 Windows에서만 됩니다.", file=sys.stderr)
        return 1
    cfg = cfgmod.load_config()
    src = Path(__file__).resolve().parent.parent / "taskpane" / "manifest.xml"
    if not src.exists():
        print(f"매니페스트를 찾을 수 없습니다: {src}", file=sys.stderr)
        return 1

    # 인증서 생성 + 신뢰(대화형 데스크톱 세션이면 성공. WebView2가 로컬
    # HTTPS를 경고 없이 로드하려면 필요하다).
    certs.ensure_dev_certificate(cfgmod.config_dir())

    # 매니페스트의 포트를 실제 설정 포트로 맞춰 로컬에 설치.
    text = src.read_text(encoding="utf-8").replace(
        "127.0.0.1:51837", f"127.0.0.1:{cfg.taskpane_port}"
    )
    dst_dir = cfgmod.config_dir() / "excel-addin"
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / "manifest.xml"
    dst.write_text(text, encoding="utf-8")

    # 개발자 사이드로드 레지스트리(값 이름=값 데이터=매니페스트 경로).
    result = subprocess.run(
        ["reg", "add", _WEF_DEVELOPER_KEY, "/v", str(dst), "/t", "REG_SZ", "/d", str(dst), "/f"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"레지스트리 등록 실패: {result.stderr.strip()}", file=sys.stderr)
        return 1
    print("Excel 리본 애드인 설치 완료.")
    print("Excel을 껐다 다시 켜면 홈 탭 오른쪽에 'XGEN > xgen-seepage 채팅' 버튼이 뜹니다.")
    print("버튼을 누르면 채팅 패널이 열립니다(`xgen-seepage run`이 켜져 있어야 함).")
    print("제거: `xgen-seepage uninstall-excel-addin`")
    return 0


async def cmd_uninstall_excel_addin(_args: argparse.Namespace) -> int:
    if platform.system() != "Windows":
        print("Windows에서만 됩니다.", file=sys.stderr)
        return 1
    dst = cfgmod.config_dir() / "excel-addin" / "manifest.xml"
    subprocess.run(["reg", "delete", _WEF_DEVELOPER_KEY, "/v", str(dst), "/f"], capture_output=True)
    print("Excel 리본 애드인을 제거했습니다(Excel 재시작 후 반영).")
    return 0


async def cmd_panel(_args: argparse.Namespace) -> int:
    """채팅 패널을 기본 브라우저로 연다. Excel 리본 버튼(Office 애드인)이
    이 Office 버전에서 안 뜨더라도 패널을 쓸 수 있는 확실한 경로다 - 패널은
    office.js에 의존하지 않는 순수 웹 UI라 브라우저로 열어도 그대로 동작한다.
    `xgen-seepage run`이 켜져 있어야 한다(패널은 그 로컬 서버에 붙는다)."""
    cfg = cfgmod.load_config()
    url = f"https://127.0.0.1:{cfg.taskpane_port}/index.html"
    print(f"채팅 패널: {url}")
    print("(`xgen-seepage run`이 켜져 있어야 합니다. 자체서명 인증서 경고가 뜨면 최초 1회 신뢰하세요.)")
    try:
        webbrowser.open(url)
    except Exception as e:
        print(f"브라우저 자동 실행 실패({e}). 위 URL을 직접 여세요.", file=sys.stderr)
    return 0


async def _ensure_valid_token(cfg: cfgmod.SeepageConfig, auth: AuthApi) -> str | None:
    """유효한 access_token을 반환한다. 만료됐으면 refresh_token으로 재발급
    시도 후 키체인을 갱신한다. 둘 다 실패하면 None(재로그인 필요)."""
    access = cfgmod.get_access_token(cfg.server_url)
    if not access:
        return None
    refresh_token = cfgmod.get_refresh_token(cfg.server_url)
    try:
        user, new_access = await auth.validate(access, refresh_token)
    except ApiError:
        return None
    if user is not None:
        return access
    if new_access:
        cfgmod.set_tokens(cfg.server_url, new_access, refresh_token)
        return new_access
    if refresh_token:
        try:
            refreshed = await auth.refresh(refresh_token)
        except ApiError:
            refreshed = None
        if refreshed:
            cfgmod.set_tokens(cfg.server_url, refreshed, refresh_token)
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
            port=cfg.taskpane_port,
            cert_path=cert_path,
            key_path=key_path,
            xgen_server_url=cfg.server_url,
            allow_private_certificate=cfg.allow_private_certificate,
            get_token=get_token,
            get_chat_workflow_id=lambda: cfg.chat_workflow_id,
            get_server_info=lambda: (cfg.server_url, cfg.username or ""),
        )
        tasks.append(asyncio.create_task(taskpane.run()))
        panel_url = f"https://127.0.0.1:{taskpane.port}/index.html"
        print(f"채팅 패널: {panel_url}")
        print("  (Excel 리본에 버튼이 안 뜨는 Office라도 이 URL을 브라우저로 열면 그대로 쓸 수 있습니다. `xgen-seepage panel`로도 열림.)")
        if args.open_panel:
            try:
                webbrowser.open(panel_url)
            except Exception:
                pass

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
        if taskpane is not None:
            await taskpane.aclose()
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
    run_p.add_argument(
        "--open-panel",
        action="store_true",
        default=False,
        help="시작하면서 채팅 패널을 기본 브라우저로 자동으로 연다",
    )
    run_p.set_defaults(func=cmd_run)

    panel_p = sub.add_parser("panel", help="채팅 패널을 기본 브라우저로 연다(run이 켜져 있어야 함)")
    panel_p.set_defaults(func=cmd_panel)

    install_p = sub.add_parser(
        "install-excel-addin",
        help="Excel 리본에 XGEN 채팅 버튼을 설치한다(폐쇄망 OK, 클라우드 불필요)",
    )
    install_p.set_defaults(func=cmd_install_excel_addin)

    uninstall_p = sub.add_parser("uninstall-excel-addin", help="Excel 리본 XGEN 버튼 제거")
    uninstall_p.set_defaults(func=cmd_uninstall_excel_addin)

    chat_wf_p = sub.add_parser(
        "chat-workflow", help="Excel 태스크팬 채팅을 연결할 XGEN 워크플로우 관리"
    )
    chat_wf_sub = chat_wf_p.add_subparsers(dest="chat_workflow_command", required=True)

    chat_wf_list_p = chat_wf_sub.add_parser("list", help="연결 가능한 워크플로우 목록")
    chat_wf_list_p.set_defaults(func=cmd_chat_workflow_list)

    chat_wf_set_p = chat_wf_sub.add_parser("set", help="태스크팬 채팅을 특정 워크플로우에 연결")
    chat_wf_set_p.add_argument("workflow_id", help="`chat-workflow list`에서 확인한 workflow_id")
    chat_wf_set_p.set_defaults(func=cmd_chat_workflow_set)

    server_p = sub.add_parser("server", help="활성 XGEN 서버(jeju/dev/prod 등) 확인/전환")
    server_sub = server_p.add_subparsers(dest="server_command", required=True)

    server_list_p = server_sub.add_parser("list", help="로그인해 본 서버 목록과 토큰 보유 여부")
    server_list_p.set_defaults(func=cmd_server_list)

    server_use_p = server_sub.add_parser("use", help="활성 서버를 바꾼다(토큰 있으면 재로그인 불필요)")
    server_use_p.add_argument("server", help="서버 URL 또는 `server list`의 번호")
    server_use_p.set_defaults(func=cmd_server_use)

    return p


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = build_parser()
    args = parser.parse_args()
    return asyncio.run(args.func(args))


if __name__ == "__main__":
    sys.exit(main())
