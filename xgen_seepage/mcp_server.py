"""MCP stdio server. XGEN Connector의 로컬 MCP, 또는 Claude Desktop/Code에서
xgen-seepage 도구를 호출하게 해준다.

실행:
    python -m xgen_seepage.mcp_server
    # 또는 설치 후
    xgen-seepage-mcp

XGEN Connector 설정 (Settings → 로컬 MCP)에 command로 위 실행법을 등록하면,
연결된 XGEN 에이전트(agent_xgen/agent_harness/agent_geny) 세션에 이 도구들이
자동 주입된다. 별도 워크플로우 편집이 필요 없다. 자세한 연동 절차는 저장소
`examples/xgen_connector_local_mcp.md` 참조.

Claude Desktop 설정 예시 (~/Library/Application Support/Claude/claude_desktop_config.json):
{
  "mcpServers": {
    "xgen-seepage": {
      "command": "/path/to/venv/bin/python",
      "args": ["-m", "xgen_seepage.mcp_server"]
    }
  }
}
"""
from __future__ import annotations

import asyncio
import json
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from .tools import TOOL_DEFINITIONS, call_tool

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("xgen-seepage-mcp")

server: Server = Server("xgen-seepage")


@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(name=t["name"], description=t["description"], inputSchema=t["input_schema"])
        for t in TOOL_DEFINITIONS
    ]


@server.call_tool()
async def on_call_tool(name: str, arguments: dict) -> list[TextContent]:
    log.info("tool call: %s %s", name, list(arguments.keys()))
    result = call_tool(name, arguments)
    return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]


async def main() -> None:
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


def main_sync() -> None:
    """Console script entry point."""
    asyncio.run(main())


if __name__ == "__main__":
    main_sync()
