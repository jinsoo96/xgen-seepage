# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)
[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

XGEN 에이전트가 사용자의 Excel/CSV 파일을 직접 읽고 편집하도록 연결하는 커넥터입니다.
XGEN 서버에 로그인해 붙는 단일 프로세스로 동작하며, 지금 화면에 열려 있는 통합문서를
저장 없이 그 자리에서 수정합니다.

다루는 형식은 `.xlsx` 와 `.csv` 입니다.

| 형식 | 방식 |
|---|---|
| `.xlsx` (열려 있는 Excel) | 실행 중인 Excel 프로세스에 붙어 실시간으로 읽고 씁니다(수식 포함). |
| `.xlsx` (Excel 미설치) | LibreOffice(UNO)로 파일을 직접 열어 같은 셀 입출력을 제공합니다. |
| `.csv` | 인코딩·구분자를 자동 감지해 읽고 편집합니다. 필요하면 Excel로 올려 이어 씁니다. |

## 지원 환경

| 대상 | 필요 조건 |
|---|---|
| 열려 있는 Excel 실시간 편집 | Windows 또는 macOS + 로컬 Excel |
| Excel 없이 xlsx 편집 | 임의 OS + LibreOffice |
| CSV 편집 | 임의 OS |

## 설치

### 1. 폐쇄망 · 인터넷이 없는 환경

대상 머신에 별도 런타임을 설치하지 않고, 미리 빌드한 단일 실행파일 하나만 반입해
사용합니다.

1. [Releases](https://github.com/jinsoo96/xgen-seepage/releases) 에서 대상 OS에 맞는
   실행파일을 내려받습니다.

   | OS | 파일 |
   |---|---|
   | Windows | `xgen-seepage-connector.exe` |
   | macOS | `xgen-seepage-connector` |

2. 대상 머신으로 복사한 뒤 그대로 실행합니다(`login` / `run` / `status` /
   `install-excel-addin` 동일).

Releases에 올릴 실행파일을 직접 만들려면, 인터넷이 되는 빌드 머신에서 다음을
실행합니다. PyInstaller는 크로스컴파일을 지원하지 않으므로 배포 대상과 같은 OS에서
빌드해야 합니다.

```bash
git clone https://github.com/jinsoo96/xgen-seepage.git
cd xgen-seepage
pip install -e ".[live,build]"
python -m PyInstaller packaging/xgen-seepage-connector.spec
# 산출물: dist/xgen-seepage-connector(.exe)
```

### 2. Python 환경 (온라인 · 개발)

```bash
# 열려 있는 통합문서까지 실시간 편집: [live]
pip install "xgen-seepage[live] @ git+https://github.com/jinsoo96/xgen-seepage.git"

# CSV/LibreOffice만 사용: [live] 없이
pip install "xgen-seepage @ git+https://github.com/jinsoo96/xgen-seepage.git"
```

## 빠른 시작

```bash
xgen-seepage login    # XGEN 서버 URL과 계정으로 로그인 (최초 1회)
xgen-seepage run      # 에이전트-도구 브릿지와 채팅 패널 상주 (Ctrl+C로 종료)
```

- `login`은 **한 번만** 하면 됩니다. 토큰이 저장돼(OS 키체인 또는 파일) 이후 `run`·재부팅마다
  자동으로 그 토큰을 씁니다. 매번 로그인하지 않습니다.
- 엑셀을 편집하는 동안 `run`은 **그 머신에서 계속 떠 있어야** 합니다. 실제 셀 편집이 이
  로컬 프로세스에서 일어나기 때문입니다(아래 [아키텍처](#아키텍처와-실행-파이프라인) 참고).

`run`이 떠 있으면 XGEN에서 에이전트와 채팅할 때 그 에이전트가 이 머신의 Excel/CSV를
도구로 사용합니다. Excel 안에서 바로 쓰려면 아래 중 하나를 실행합니다.

```bash
xgen-seepage install-excel-addin   # Excel 홈 탭에 채팅 버튼 설치 (uninstall-excel-addin으로 제거)
xgen-seepage panel                 # 채팅 패널을 기본 브라우저로 열기
```

패널에서 에이전트(워크플로우)를 선택해 채팅하면 그 에이전트가 열린 Excel을 편집합니다.
사용할 LLM 모델은 해당 에이전트플로우의 노드 설정을 따릅니다. 노드에 모델이 설정돼
있지 않으면 패널이 안내 메시지를 표시합니다.

## 배포 (폐쇄망 · 여러 사용자)

엑셀 편집은 사용자 PC에서 일어나므로, **커넥터는 엑셀을 쓸 각 사용자 머신에서 떠 있어야
합니다.** 서버가 남의 PC에 열린 엑셀 안으로 들어갈 방법이 없기 때문이며, 로컬 컴포넌트가
필요한 것은 스프레드시트를 다루는 AI 애드인 전반의 공통점입니다.

사용자가 직접 설치·실행하게 하지 않고, 관리자가 다음을 한 번 준비하면 **사용자는 부팅만
해도 됩니다**:

1. **중앙 배포** — 단일 실행파일(exe/바이너리)을 GPO·SCCM·MDM이나 표준 이미지로 전 PC에
   배포합니다. 파이썬·인터넷이 필요 없어 폐쇄망 반입에 적합합니다.
2. **부팅 자동실행** — `xgen-seepage run`을 OS 로그온 시 자동 실행되도록 등록합니다.
   - Windows: 시작프로그램 폴더에 바로가기, 또는 작업 스케줄러(트리거: 로그온).
   - macOS: `~/Library/LaunchAgents`에 LaunchAgent plist 등록 후 `launchctl load`.
3. **토큰 provisioning** — 최초 1회 `login`으로 토큰을 심어두거나, 배포 스크립트가 환경변수
   `XGEN_SEEPAGE_SERVER_URL` / `XGEN_SEEPAGE_EMAIL` / `XGEN_SEEPAGE_PASSWORD`로 자동
   로그인하게 합니다. 이후 재부팅마다 저장된 토큰으로 자동 접속합니다.

→ 이 셋이면 사용자는 부팅만 해도 커넥터가 이미 연결돼 있어, XGEN에서 에이전트와 채팅하면
그 에이전트가 곧바로 엑셀을 편집합니다. 커넥터는 한 번 뜨면 연결이 끊겨도 자동으로
재접속합니다(지수 백오프).

**에이전트가 도구를 받는 방식** — 커넥터가 연결돼 있으면, XGEN의 표준 에이전트 노드
(harness · xgen · geny)가 이 도구들을 **실행 시점에 자동으로 주입**받습니다. 그래프에 별도
노드를 넣을 필요가 없습니다. 즉 "엑셀 환경에 연결 → 그 계정의 어떤 에이전트든 엑셀 도구를
사용"이 성립합니다.

## 아키텍처와 실행 파이프라인

xgen-seepage는 XGEN 서버가 도구 호출을 이 로컬 프로세스로 되돌려 보내고, 이 프로세스가
로컬 Excel/LibreOffice/CSV에 그 호출을 적용하는 역방향 브릿지 구조입니다.

```
 XGEN 서버 ────────────────────────────  이 머신 (xgen-seepage run)
 ┌───────────────────┐                    ┌───────────────────────────────┐
 │ 에이전트 실행       │  ① 로그인(HTTPS)   │ connector/                    │
 │ (웹 UI · 패널)      │ ◀───────────────  │  - auth: 토큰 발급/검증        │
 │                   │                    │  - config: 서버별 토큰 저장    │
 │ ② tool_call        │  WebSocket 브릿지  │  - bridge: 카탈로그 광고/실행  │
 │   ↓               │ ◀───────────────▶ │                               │
 │ ③ mcp_call ───────┼───────────────────┼─▶ tools.call_tool             │
 │ ④ mcp_result ◀────┼───────────────────┼── ├─ live_adapter (xlwings)    │
 │                   │                    │   ├─ libreoffice_adapter(UNO) │
 └───────────────────┘                    │   └─ csv_adapter              │
                                          └───────────────┬───────────────┘
                                                          ▼
                                            열린 Excel / xlsx 파일 / CSV
```

**실행 파이프라인**

1. **로그인** — `login`이 XGEN에 인증해 access/refresh 토큰을 받습니다. 토큰은 서버
   URL별로 나눠 OS 키체인에 저장하고, 키체인에 접근할 수 없는 환경(헤드리스·서비스
   세션·폐쇄망)에서는 설정 폴더의 파일(권한 0600)로 폴백합니다.
2. **카탈로그 광고** — `run`이 XGEN의 `/api/tools/ws/connector-mcp/{user_id}` WebSocket에
   붙어, 이 프로세스가 제공하는 도구 목록을 `hello` 프레임으로 광고합니다. 서버는 이
   목록을 사용자 세션에 등록합니다.
3. **도구 실행** — 에이전트가 도구를 호출하면 서버가 `mcp_call`을 브릿지로 내려보내고,
   이 프로세스가 `tools.call_tool`로 처리해 결과를 `mcp_result`로 돌려줍니다. 실제 셀
   조작은 로컬 프로세스 간 통신으로만 일어나며, 편집 결과는 저장 없이 화면에 즉시
   반영됩니다.
4. **셀 입출력 계층** — 호출은 대상에 따라 세 어댑터 중 하나로 분기됩니다.
   - `live_adapter` — 실행 중인 Excel에 xlwings로 붙습니다. Windows는 COM, macOS는
     AppleScript로 동작하며, 수식·서식·병합을 보존한 채 편집합니다.
   - `libreoffice_adapter` — Excel이 없는 환경에서 LibreOffice(UNO)로 xlsx를 직접
     엽니다.
   - `csv_adapter` — 인코딩·구분자를 감지해 CSV를 편집하고, 원본 인코딩으로 재저장합니다.

**인프라 관점**

- **의존성** — 실행에 필요한 것은 XGEN 서버 연결뿐입니다. 별도 데스크톱 클라이언트나
  Microsoft 애드인 호스팅 인프라가 필요하지 않으며, 채팅 패널도 XGEN 로그인에만
  의존하는 웹 UI입니다.
- **TLS** — 사내 내부 CA로 발급한 HTTPS 인증서를 쓰는 XGEN이라도, 그 CA가 OS 신뢰
  저장소에 설치돼 있으면 그대로 접속합니다(`truststore`로 OS 저장소를 참조). CA가
  OS에 없는 경우에만 `login --allow-private-certificate`로 해당 서버에 한해 검증을
  완화할 수 있습니다.
- **다중 서버** — 여러 XGEN(예: 개발/스테이징/운영)을 오갈 때 토큰을 서버별로
  저장하고 `server use`로 재로그인 없이 전환합니다. 현재 붙은 서버·계정은 패널 헤더와
  `status`에서 확인합니다.

## CLI

| 명령 | 설명 |
|---|---|
| `xgen-seepage login` | XGEN 서버에 로그인하고 토큰을 서버별로 저장합니다. |
| `xgen-seepage run` | 에이전트-도구 브릿지와 채팅 패널 서버를 상주시킵니다(`--open-panel`로 패널 자동 열기). |
| `xgen-seepage server list` / `use` | 로그인해 본 서버를 확인하고 전환합니다. |
| `xgen-seepage panel` | 채팅 패널을 기본 브라우저로 엽니다(`run` 필요). |
| `xgen-seepage install-excel-addin` | Excel 리본에 채팅 버튼을 설치합니다(`uninstall-excel-addin`으로 제거). |
| `xgen-seepage chat-workflow list` / `set` | 패널 기본 에이전트(워크플로우)를 관리합니다. |
| `xgen-seepage status` | 현재 설정·토큰 유효성·권한을 확인합니다. |
| `xgen-seepage logout` | 저장된 토큰을 삭제합니다. |

## 파이썬 API

```python
from xgen_seepage import live_adapter, csv_adapter

books = live_adapter.list_open_workbooks()
wb_id = books[0].workbook_id

schema = live_adapter.get_sheet_schema(wb_id, sheet=0)   # 레이아웃·병합·수식 존재 여부
live_adapter.set_cell(wb_id, sheet=0, row=1, col=2, value="1500")
live_adapter.set_cell(wb_id, sheet=0, row=1, col=3, value="=C2*1.1", as_formula=True)
data = live_adapter.read_range(wb_id, sheet=0, row0=0, col0=0, row1=10, col1=5)

# CSV: 인코딩/구분자 자동 감지, 필요 시 Excel로 승격
csv_adapter.set_cell("sales.csv", row=1, col=1, value="9800")
wb = csv_adapter.open_in_excel("sales.csv")

# Excel 미설치 환경: LibreOffice로 xlsx 직접 열기
from xgen_seepage import libreoffice_adapter as lo
doc = lo.open_document("report.xlsx")["path"]
lo.set_cell(doc, sheet=0, row=1, col=2, value="1500")
lo.save(doc)
```

## 제공 도구

`run`이 아래 도구를 XGEN 에이전트 세션에 광고합니다.

**조회** — 통합문서/시트 구조 파악, 셀·범위 읽기, 표 범위 인식

**셀·수식 편집** — 셀/범위 쓰기, 행 추가, 찾기·바꾸기, 수식 재계산

**구조 편집** — 행·열 삽입/삭제/숨김, 범위 복사, 정렬, 자동필터(Windows)

**시트·이름** — 시트 추가/이름변경/삭제/이동, 이름 정의 관리

**서식** — 배경색·조건부 행 강조·글꼴·표시형식·테두리·자동 줄바꿈·병합·틀 고정·
열너비/행높이·유효성 목록

**LibreOffice · CSV** — Excel 미설치 환경의 xlsx 편집, CSV 조회/편집/생성

전체 도구 목록과 시그니처는 [`xgen_seepage/tools.py`](xgen_seepage/tools.py)에서
확인할 수 있습니다.

## stdio MCP 서버

기본 경로는 `xgen-seepage run`이지만, 로컬(stdio) MCP를 지원하는 클라이언트에 얹어
쓰려면 stdio MCP 서버도 제공합니다. [`examples/local_mcp_stdio.md`](examples/local_mcp_stdio.md)를
참고하세요.

```bash
xgen-seepage-mcp        # 또는 python -m xgen_seepage.mcp_server
```

## 개발

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e ".[live,dev]"
ruff check xgen_seepage/
mypy xgen_seepage/
```

## 라이선스

Apache License 2.0. [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE)를 참고하세요.
