"""수동 검증 스크립트(테스트 스위트에 포함 안 함) — PyInstaller로 얼린
xgen-seepage-connector.exe를 실제 서브프로세스로 띄워 가짜 XGEN 서버에
login + run(WS 브릿지)까지 진짜로 시켜본다. venv 소스가 아니라 **얼린
바이너리 자체**가 로그인하고 hello 프레임을 보내는지 확인하는 최종
검증용. 실행 후 실제 Windows Credential Manager에 쓴 토큰은 스스로 지운다.
"""
from __future__ import annotations

import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from aiohttp import web

EXE = Path(__file__).resolve().parent / "dist" / "xgen-seepage-connector.exe"
PASSWORD = "frozen-smoke-test-pw"


def _sha256_hex(s: str) -> str:
    import hashlib
    return hashlib.sha256(s.encode()).hexdigest()


async def handle_login(request: web.Request) -> web.Response:
    body = await request.json()
    if body.get("password") == _sha256_hex(PASSWORD):
        return web.json_response({
            "success": True, "access_token": "frozen-tok", "refresh_token": "frozen-ref",
            "token_type": "bearer", "user_id": "frozen-user", "username": body["email"],
        })
    return web.json_response({"success": False, "message": "bad", "access_token": None})


hello_received = asyncio.Event()
hello_payload: dict = {}


async def handle_ws(request: web.Request) -> web.WebSocketResponse:
    auth = request.headers.get("Authorization")
    ws = web.WebSocketResponse()
    await ws.prepare(request)
    raw = await ws.receive_str()
    msg = json.loads(raw)
    hello_payload["msg"] = msg
    hello_payload["auth_header"] = auth
    hello_received.set()
    await ws.send_str(json.dumps({"type": "ready", "catalog_id": msg["catalog_id"], "tool_count": len(msg["tools"])}))
    await asyncio.sleep(0.3)
    await ws.close()
    return ws


async def main() -> int:
    if not EXE.exists():
        print(f"frozen exe not found at {EXE} — build it first", file=sys.stderr)
        return 1

    app = web.Application()
    app.router.add_post("/api/auth/login", handle_login)
    app.router.add_get("/api/tools/ws/connector-mcp/{user_id}", handle_ws)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "127.0.0.1", 0)
    await site.start()
    port = site._server.sockets[0].getsockname()[1]
    base_url = f"http://127.0.0.1:{port}"
    print(f"mock XGEN server at {base_url}")

    env = dict(os.environ)
    env["XGEN_SEEPAGE_SERVER_URL"] = base_url
    env["XGEN_SEEPAGE_EMAIL"] = "smoke@example.com"
    env["XGEN_SEEPAGE_PASSWORD"] = PASSWORD

    print("== running frozen exe: login ==")
    login_proc = subprocess.run([str(EXE), "login"], env=env, capture_output=True, text=True, timeout=30)
    print(login_proc.stdout)
    print(login_proc.stderr, file=sys.stderr)
    if login_proc.returncode != 0:
        print("FROZEN EXE LOGIN FAILED")
        await runner.cleanup()
        return 1

    print("== running frozen exe: run (bridge) — will wait for hello then kill it ==")
    run_proc = subprocess.Popen([str(EXE), "run"], env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    try:
        await asyncio.wait_for(hello_received.wait(), timeout=15)
        ok = True
    except asyncio.TimeoutError:
        ok = False
    finally:
        run_proc.terminate()
        try:
            run_proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            run_proc.kill()

    await runner.cleanup()

    print("== cleaning up: logout (removes real keychain entries this smoke test wrote) ==")
    subprocess.run([str(EXE), "logout"], env=env, capture_output=True, text=True, timeout=15)

    if not ok:
        print("FROZEN EXE NEVER SENT hello OVER THE WS BRIDGE — FAIL")
        return 1

    msg = hello_payload["msg"]
    print(f"hello received. auth_header={hello_payload['auth_header']!r}")
    print(f"catalog_id={msg['catalog_id']!r} tool_count={len(msg['tools'])}")
    tool_names = sorted(t["name"] for t in msg["tools"])
    print(f"advertised tools: {tool_names}")
    assert hello_payload["auth_header"] == "Bearer frozen-tok", "wrong bearer token sent by frozen exe"
    assert "inspect_csv" in tool_names, "expected tool missing from frozen exe's catalog"
    print("PASS: frozen exe logged in and bridged its real tool catalog over a live WS connection.")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
