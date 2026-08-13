# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)

**오프라인 설치형 XGEN Excel/CSV 커넥터.** 챗지피티나 클로드가 엑셀 안에
플러그인처럼 붙어서 편집·붙여넣기하는 것의 XGEN 버전 — 단, **인터넷 연결
없이 로컬에서만** 동작한다.

지금 로컬 Excel에서 **열려 있는** 통합문서를 실시간으로 읽고 쓴다(수식 포함,
파일 저장 불필요 — 사용자가 화면에서 바로 결과를 본다). CSV는 인코딩/구분자를
자동 감지해 읽고 편집하며, 필요하면 Excel로 올려 같은 라이브 경로로 넘어간다.

- 📦 [ARCHITECTURE.md](ARCHITECTURE.md) — Claude for Excel / ChatGPT for Excel이
  실제로 어떻게 동작하는지 조사한 내용과, 그걸 오프라인 제약 하에서 어떻게
  재설계했는지의 근거 전부.
- 🔌 [examples/xgen_connector_local_mcp.md](examples/xgen_connector_local_mcp.md) —
  `xgen-connector`의 "로컬 MCP"에 등록해 XGEN 에이전트와 바로 연결하는 법.

## 왜 태스크팬이 아니라 xlwings인가

Claude for Excel/ChatGPT for Excel은 실제로는 **Office.js 태스크팬 애드인**이다
— Excel 옆 패널이 Excel JavaScript API로 셀을 읽고 쓴다. 이 플랫폼 자체는
`localhost` 인증서로 오프라인도 가능하지만, 매니페스트 사이드로딩·인증서
신뢰 설정 등 설치 마찰이 크다.

xgen-seepage는 대신 **xlwings**(로컬 Excel 프로세스에 COM/AppleScript로 직접
붙는 라이브러리)를 쓴다. 웹서버도 인증서도 매니페스트도 없다 — `pip install`
만으로 끝나는 순수 로컬 프로세스 통신이라 인터넷 요구사항이 원천적으로 없다.
자세한 비교와 근거는 [ARCHITECTURE.md](ARCHITECTURE.md) 참조.

## 설치

```bash
pip install xgen-seepage[live]   # 라이브(열려 있는 통합문서) 편집 포함
# CSV만 필요하면
pip install xgen-seepage
```

`[live]`는 Windows 또는 macOS에서 로컬 Excel이 설치돼 있어야 동작한다.

## 빠른 시작 — 파이썬 API

```python
from xgen_seepage import live_adapter, csv_adapter

# 지금 열려 있는 모든 통합문서 조회
books = live_adapter.list_open_workbooks()
wb_id = books[0].workbook_id

# 시트 구조 파악 (레이아웃, 병합 셀, 수식 존재 여부)
schema = live_adapter.get_sheet_schema(wb_id, sheet=0)
print(schema.preview)

# 셀 하나 실시간으로 갱신 — 저장 없이 즉시 화면에 반영됨
live_adapter.set_cell(wb_id, sheet=0, row=1, col=2, value="1500")

# 수식도 그대로 쓴다
live_adapter.set_cell(wb_id, sheet=0, row=1, col=3, value="=C2*1.1", as_formula=True)

# 범위 벌크 읽기 (값 + 수식 둘 다)
data = live_adapter.read_range(wb_id, sheet=0, row0=0, col0=0, row1=10, col1=5)
```

```python
# CSV: 인코딩/구분자 자동 감지
table = csv_adapter.load_table("sales_2026.csv")
print(table.encoding, table.header)

csv_adapter.set_cell("sales_2026.csv", row=1, col=1, value="9800")

# 수식·서식이 필요하면 Excel로 승격
wb = csv_adapter.open_in_excel("sales_2026.csv")
live_adapter.set_cell(wb["workbook_id"], sheet=0, row=0, col=3, value="=B1*1.1", as_formula=True)
```

## MCP 서버로 사용 — XGEN Connector / Claude Desktop / Claude Code

```bash
xgen-seepage-mcp
# 또는
python -m xgen_seepage.mcp_server
```

XGEN Connector의 "로컬 MCP" 설정에 위 명령을 등록하면, 연결된 XGEN 에이전트
(`agent_xgen`/`agent_harness`/`agent_geny`) 세션에 아래 도구가 자동 주입된다
— 별도 워크플로우 편집이 필요 없다. 절차는
[examples/xgen_connector_local_mcp.md](examples/xgen_connector_local_mcp.md).

Claude Desktop 설정(`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "xgen-seepage": {
      "command": "/absolute/path/to/venv/bin/python",
      "args": ["-m", "xgen_seepage.mcp_server"]
    }
  }
}
```

## 노출되는 도구 (14개)

| 도구 | 설명 |
|---|---|
| `list_open_workbooks` | 로컬에서 실행 중인 모든 Excel의 열린 통합문서 목록. **항상 첫 호출** |
| `get_live_schema` | 시트 크기·미리보기·병합·수식 존재 여부 |
| `get_live_cell` | 셀 하나의 값+수식+병합 정보 전체(절단 없이) |
| `set_live_cell` | 셀 하나 즉시 갱신(수식 지원, 숫자 자동판정, 병합 anchor 리다이렉트) |
| `read_live_range` | 범위 값/수식 벌크 읽기(최대 50,000셀) |
| `write_live_range` | 범위 벌크 쓰기(붙여넣기와 동일 동작) |
| `append_live_row` | 사용 범위 아래에 행 추가 |
| `activate_live_cell` | Excel 창을 앞으로 + 셀 선택(에이전트가 뭘 보는지 사용자가 눈으로 확인) |
| `inspect_csv` | CSV 구조 조사(인코딩/구분자 자동 감지 + 미리보기) |
| `get_csv_cell` / `set_csv_cell` | CSV 셀 조회/편집(원본 인코딩·구분자 보존 재저장) |
| `append_csv_row` | CSV 끝에 행 추가 |
| `write_csv_table` | CSV 신규 생성/통째 덮어쓰기 |
| `open_csv_in_excel` | CSV를 로컬 Excel로 열어 라이브 도구로 이어서 편집 |

## PlateerLab 생태계와의 관계

이 프로젝트는 처음부터 새로 만들지 않았다. 회사 GitHub의 기존 레포 세 개를
근거로 삼았다(자세한 재사용 내역은 `NOTICE`, 설계 근거는 `ARCHITECTURE.md` §3):

- **[document-adapter](https://github.com/PlateerLab/document-adapter)** — 닫힌
  xlsx 파일 편집(병합 셀·수식·MCP 노출)의 원본. xgen-seepage는 이걸 다시
  구현하지 않고, **셀 편집 시맨틱만 이식**해 "지금 열려 있는 통합문서"라는
  새 표면에 적용했다. 닫힌 파일 편집이 필요하면 document-adapter를 그대로
  같이 등록해서 쓰면 된다.
- **[xgen-doc2chunk](https://github.com/PlateerLab/xgen-doc2chunk)** — CSV
  인코딩 자동 감지 전략(BOM→chardet→후보목록→latin-1)의 출처.
- **[xgen-connector](https://github.com/PlateerLab/xgen-connector)** — 이 MCP
  서버를 실제 XGEN 에이전트 세션에 연결해 주는 "로컬 MCP" 브릿지. 채팅 UI,
  인증, 에이전트 연결은 전부 여기서 이미 해결된 문제라 xgen-seepage는 새로
  만들지 않았다.

## 개발

```bash
python -m venv .venv
.venv/Scripts/activate       # Windows
pip install -e ".[live,dev]"
pytest tests/ -v
ruff check xgen_seepage/ tests/
```

`tests/test_csv_adapter.py`는 실제 인코딩(BOM/CP949)·구분자 자동 감지, 셀
편집 왕복, MCP dispatcher 경로를 검증한다(Excel 없이도 전부 동작). `tests/
test_live_adapter.py`는 xlwings는 설치돼 있지만 실행 중인 Excel이 없는
환경에서 라이브 도구들이 명확한 에러(`WorkbookNotFoundError`/
`ExcelUnavailableError`)로 우아하게 실패하는지 검증한다 — 실제로 열린
통합문서의 셀 IO 자체는 Excel이 설치된 머신에서 별도 검증이 필요하다
(`ARCHITECTURE.md` §6 참조).

## 라이선스

**Apache License 2.0** — [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조. 이
소프트웨어나 그 파생물을 재배포할 때는 저작권·라이선스 고지와 `NOTICE`
파일의 출처 표시를 유지해야 합니다(라이선스 §4).
