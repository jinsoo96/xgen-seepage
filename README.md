# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)

**폐쇄망에서도 설치되는 독립형 XGEN Excel/CSV 커넥터.** 챗지피티나 클로드가
엑셀 안에 플러그인처럼 붙어서 편집·붙여넣기하는 것의 XGEN 버전 — 단, XGEN
서버에 직접 로그인해서 붙는 **독립 실행 프로그램**이고(다른 앱 설치 불필요),
인터넷 없이도 설치·운영 가능하다.

- **로그인 한 번**(`xgen-seepage login`) → XGEN 서버·계정에 붙는다.
- **`xgen-seepage run`을 상주**시키면, 사용자가 XGEN 어디서 어떤 에이전트와
  채팅하든 그 에이전트가 사용자 로컬 Excel/CSV를 도구로 쓸 수 있게 된다.
- 지금 로컬 Excel에서 **열려 있는** 통합문서를 실시간으로 읽고 쓴다(수식 포함,
  파일 저장 불필요 — 사용자가 화면에서 바로 결과를 본다). CSV는 인코딩/구분자를
  자동 감지해 읽고 편집하며, 필요하면 Excel로 올려 같은 라이브 경로로 넘어간다.

- 📦 [ARCHITECTURE.md](ARCHITECTURE.md) — Claude for Excel / ChatGPT for Excel이
  실제로 어떻게 동작하는지 조사한 내용, XGEN 서버 로그인·에이전트-도구 브릿지의
  실제 와이어 프로토콜, PyInstaller 단일 실행파일 패키징과 실제로 찾아 고친
  버그까지 전부.

## 왜 독립 실행 프로그램인가

Claude for Excel/ChatGPT for Excel은 실제로는 **Office.js 태스크팬 애드인**이다
— Excel 옆 패널이 Excel JavaScript API로 셀을 읽고 쓴다. 이 플랫폼 자체는
`localhost` 인증서로 오프라인도 가능하지만, 매니페스트 사이드로딩·인증서
신뢰 설정 등 설치 마찰이 크다.

xgen-seepage는 셀 IO에 **xlwings**(로컬 Excel 프로세스에 COM/AppleScript로
직접 붙는 라이브러리)를 쓰고, XGEN 서버 연결에는 **자체 구현한 로그인+WebSocket
브릿지**를 쓴다 — `xgen-connector`(PlateerLab의 다른 XGEN 클라이언트)가 설치돼
있든 없든, 그 앱의 설정을 전혀 건드리지 않고 독립적으로 동작한다. 자세한 비교와
근거는 [ARCHITECTURE.md](ARCHITECTURE.md) 참조.

## 설치 & 사용

**개발/테스트용 (인터넷 있는 환경):**

```bash
pip install "xgen-seepage[live]"   # 라이브(열려 있는 통합문서) 편집 포함
xgen-seepage login                 # XGEN 서버 URL + 계정 (대화형 프롬프트)
xgen-seepage run                   # 브릿지 상주 시작 (Ctrl+C로 종료)
```

**폐쇄망용 (단일 설치파일):** `packaging/` 아래에 PyInstaller 빌드 레시피가
있다 — 인터넷 있는 머신에서 미리 얼려서 `xgen-seepage-connector.exe` 하나로
만든 뒤, 그 실행파일만 폐쇄망에 반입한다.

```powershell
pip install -e ".[dev]"
python -m PyInstaller --onefile --name xgen-seepage-connector `
  --collect-all keyring packaging/run_connector.py
```

빌드 산출물(`packaging/dist/xgen-seepage-connector.exe`)은 어떤 파이썬도
설치돼 있지 않은 Windows 머신에서 그대로 실행된다 — `.exe login` /
`.exe run` / `.exe status` / `.exe logout`. 실제로 얼려서 검증한 내용과
찾아 고친 버그는 `ARCHITECTURE.md` §7 참조.

**비대화형 배포**(대량 설치 시 로그인을 스크립트로): 환경변수
`XGEN_SEEPAGE_SERVER_URL` / `XGEN_SEEPAGE_EMAIL` / `XGEN_SEEPAGE_PASSWORD`로
`login`을 자동화할 수 있다.

## CLI

| 명령 | 설명 |
|---|---|
| `xgen-seepage login` | XGEN 서버에 로그인, 토큰을 OS 키체인에 저장 |
| `xgen-seepage run` | 에이전트-도구 브릿지를 상주시킨다(포그라운드) |
| `xgen-seepage status` | 현재 설정·토큰 유효성 확인 |
| `xgen-seepage logout` | 저장된 토큰 삭제 |

## 파이썬 API — 직접 라이브 편집

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

## 노출되는 도구 (14개 — `xgen-seepage run`이 XGEN 에이전트 세션에 자동 광고)

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

## 대안 경로 — 이미 xgen-connector나 Claude Desktop을 쓰고 있다면

기본 경로는 위의 독립 `xgen-seepage run`이지만, 이미 `xgen-connector`(또는
Claude Desktop/Code)를 쓰고 있어서 그쪽 Local MCP 설정에 그냥 얹고 싶다면
stdio MCP 서버도 남겨뒀다:

```bash
xgen-seepage-mcp
# 또는
python -m xgen_seepage.mcp_server
```

절차는 [examples/xgen_connector_local_mcp.md](examples/xgen_connector_local_mcp.md).

## PlateerLab 생태계와의 관계 — 참고만 했다, 의존하지 않는다

이 프로젝트는 회사 GitHub의 기존 레포 세 개를 **로직 참고 자료**로만 썼다.
런타임에 그 어떤 레포도 설치돼 있을 필요가 없다(자세한 근거는 `NOTICE`,
설계 판단은 `ARCHITECTURE.md` §3):

- **[document-adapter](https://github.com/PlateerLab/document-adapter)** — 닫힌
  xlsx 파일 편집(병합 셀·수식)의 원본. **셀 편집 시맨틱만** 이식해 "지금 열려
  있는 통합문서"라는 새 표면에 적용했다. 코드도, 패키지 의존성도 가져오지
  않았다.
- **[xgen-doc2chunk](https://github.com/PlateerLab/xgen-doc2chunk)** — CSV
  인코딩 자동 감지 전략(BOM→chardet→후보목록→latin-1)의 출처.
- **[xgen-connector](https://github.com/PlateerLab/xgen-connector)** — 로그인
  API(`/api/auth/*`)와 에이전트-도구 WebSocket 브릿지
  (`/api/tools/ws/connector-mcp/{user_id}`)의 와이어 프로토콜을 그 소스에서
  확인하고, **xgen-seepage 자체 코드로 독립 재구현**했다. 이 앱이 설치돼 있을
  필요가 없다.

## 개발

```bash
python -m venv .venv
.venv/Scripts/activate       # Windows
pip install -e ".[live,dev]"
pytest tests/ -v
ruff check xgen_seepage/ tests/
```

- `tests/test_csv_adapter.py` — 실제 인코딩(BOM/CP949)·구분자 자동 감지, 셀
  편집 왕복, MCP dispatcher 경로(Excel 없이도 전부 동작).
- `tests/test_live_adapter.py` — xlwings는 설치돼 있지만 실행 중인 Excel이
  없는 환경에서 라이브 도구들이 명확한 에러로 우아하게 실패하는지 검증.
- `tests/test_connector_*.py` — 로그인 해싱/설정파일/OS 키체인 왕복, 그리고
  **가짜 XGEN 서버(aiohttp)를 실제 로컬 소켓에 띄워** login→hello→ready→
  mcp_call→mcp_result 전체 왕복을 실제 코드로 검증(`test_connector_bridge_
  integration.py`).
- 실제로 열린 통합문서의 셀 IO 자체, 그리고 얼린 exe의 OS 키체인 저장이
  일반 데스크톱 세션에서 성공하는지는 Excel/대화형 세션이 있는 머신에서 별도
  검증이 필요하다(`ARCHITECTURE.md` §8).

## 라이선스

**Apache License 2.0** — [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조. 이
소프트웨어나 그 파생물을 재배포할 때는 저작권·라이선스 고지와 `NOTICE`
파일의 출처 표시를 유지해야 합니다(라이선스 §4).
