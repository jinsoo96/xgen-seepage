# xgen-seepage. 아키텍처 & 리서치 노트

> "챗지피티나 클로드가 엑셀 안에 플러그인처럼 붙어서 편집·붙여넣기하는" 것의 XGEN 버전을,
> **로컬/인터넷 없이도 설치 가능한 커넥터**로 만든다.

이 문서는 (1) Claude/ChatGPT의 실제 Excel 통합이 기술적으로 어떻게 동작하는지 조사한
내용, (2) 그걸 그대로 베끼지 않고 **오프라인 설치**라는 제약 하에서 XGEN 생태계
(PlateerLab의 기존 레포들)를 최대한 재사용해 재설계한 근거를 남긴다.

## 1. 실제 제품은 어떻게 동작하는가

### Claude for Excel (Anthropic)

Microsoft AppSource로 배포되는 **Office.js 태스크팬 애드인**이다. Excel 옆에 패널이
붙고, 그 패널의 JS가 Excel JavaScript API로 시트를 조회/수정한다. 트래픽 분석에 따르면
시트 전체를 매번 보내지 않고 **필요할 때만** `get_range_as_csv` 같은 tool-call로
범위를 당겨온다(반면 Word 통합은 문서 전체를 매 요청에 넣는다). 백엔드는 Anthropic
클라우드(`api.anthropic.com`)이고, 로그인한 유료 Claude 계정이 필요하다. **오프라인
동작이 아니다.**

### ChatGPT for Excel / 서드파티 (GPT for Excel 등)

같은 패턴(Office.js 태스크팬) + 추가로 `=GPT()` 같은 **Custom Function**으로 셀
수식처럼 호출하는 방식도 있다. 전부 클라우드 LLM 호출이 전제다.

### Office Add-ins 플랫폼 자체는 오프라인이 가능하다

두 제품이 클라우드에 묶여 있다고 해서 **플랫폼 자체**가 오프라인 불가능한 건
아니다:
- 태스크팬은 그냥 HTTPS로 호스팅되는 웹페이지다. `localhost`에 자체서명 인증서
  (`office-addin-dev-certs`)를 붙이면 Office의 내장 WebView2가 문제없이 로드한다.
  즉 **원리적으로는** 완전 로컬 웹서버 + Office.js 태스크팬으로 오프라인 커넥터를
  만들 수 있다.
- Office Store 없이도 **네트워크 공유 폴더 카탈로그**(XML 매니페스트, JSON 통합
  매니페스트는 이 경로 미지원)로 사이드로딩할 수 있다.
- 셀 API: `range.load("values,formulas,formulasR1C1")` + `context.sync()`로 값과
  수식을 분리해서 읽고, `range.values =` / `range.formulas =`로 쓴다.

**그런데 이 경로를 그대로 택하지 않은 이유**: 인증서 신뢰, 매니페스트 사이드로딩,
SharedRuntime 설정 등 설치 마찰이 여전히 크다. "인터넷 없이도 설치"라는 목표에
비해 얻는 게 적다. 태스크팬은 **UI 셸**일 뿐이고, 우리가 필요한 건 셀 IO다.

## 2. 대안: xlwings (채택)

`xlwings`는 Windows는 COM, macOS는 AppleScript로 **로컬에서 실행 중인 Excel
프로세스에 직접 붙는다.** 웹서버도, 인증서도, 매니페스트도 없다. `pip install`만
하면 끝나는 순수 로컬 프로세스 간 통신이라 **인터넷 요구사항이 원천적으로 없다.**

| | Office.js 태스크팬 (Claude/ChatGPT 실제 구현) | xlwings COM/AppleScript (이 프로젝트) |
|---|---|---|
| 설치 | 매니페스트 사이드로딩 + 인증서 신뢰 | `pip install` |
| 인터넷 | 태스크팬 자체는 불필요(localhost 가능) | 불필요 |
| UI | 시트 옆 패널(진짜 "붙어 있는" 느낌) | 없음(별도 채팅 UI 필요. 아래 §4) |
| 셀 IO | `range.values` / `range.formulas` | `Range.value` / `Range.formula`/`.formula2` |
| 플랫폼 | Windows/Mac/Web Excel | Windows/Mac 데스크톱 Excel만(Web 불가) |

값/수식 API가 거의 1:1 대응이라(`values`↔`.value`, `formulas`↔`.formula`) 편집
시맨틱을 옮기는 데 문제가 없었다. 대신 잃는 건 "시트 옆에 항상 붙어 있는 패널" UI다.
이건 아키텍처가 아니라 UI 계층의 문제이므로 §4에서 별도로 다룬다.

## 3. PlateerLab 생태계. 참고만 하고, 독립적으로 다시 짰다

**2026-08-13 사용자 정정**: 처음엔 "xgen-connector의 로컬 MCP 브릿지에 등록해서
쓰자"는 방향으로 설계했었다. 하지만 실제 요구사항은 달랐다. **폐쇄망에서 쓰는
독립 설치 프로그램**이어야 하고, document-adapter·xgen-doc2chunk·xgen-connector는
전부 **"로직을 참고만" 하는 대상**이지 런타임 의존성이 아니다. 즉 xgen-connector가
설치돼 있지 않아도, 그 앱의 로컬 MCP 설정을 손대지 않아도 동작해야 한다. 아래는
그 세 레포에서 실제로 뭘 참고했는지(코드는 안 가져오고 로직/프로토콜만 확인)다.

### `PlateerLab/document-adapter`. 셀 편집 시맨틱

닫힌 xlsx 파일을 openpyxl로 열어 병합 셀 인지, 수식 캐시값/수식 문자열 구분,
숫자-처럼-보이는-문자열 자동 숫자화까지 구현해 놓은 Apache-2.0 라이브러리다. 이
프로젝트가 새로 만든 건 "파일이 아니라 지금 열려 있는 통합문서"라는 한 가지뿐이라,
그 차이를 만드는 편집 시맨틱(병합 셀 anchor 리다이렉트, 숫자/문자 판정 휴리스틱)만
`document-adapter/xlsx_adapter.py`에서 이식했다(`xgen_seepage/_cellfmt.py`, 출처는
`NOTICE`). document-adapter 자체는 의존성으로 추가하지 않았다. 닫힌 파일 편집이
필요하면 그 라이브러리를 따로 쓰면 된다(§8 로드맵).

### `PlateerLab/xgen-doc2chunk`. CSV 인코딩 감지 전략

국내 실사용 CSV(EUC-KR/CP949 공공·사내 데이터)에서 검증된 "BOM 우선 → chardet
(신뢰도 임계값) → 후보 인코딩 순차 시도 → latin-1 최종 폴백" 순서를 이식했다
(`xgen_seepage/csv_adapter.py::detect_encoding`, 출처는 `NOTICE`).

### `PlateerLab/xgen-connector`. 와이어 프로토콜의 참고 자료 (의존성 아님)

`xgen-connector`(Electron 데스크톱 앱)는 XGEN 서버에 로그인하고, 로컬에 등록한
MCP 서버의 도구를 에이전트 세션에 자동 주입하는 "로컬 MCP 브릿지"를 이미 갖고
있다. **이 프로젝트는 그 브릿지에 올라타지 않는다**. 대신 소스를 읽어서 아래
와이어 프로토콜을 확인하고, xgen-seepage 자체 코드로 **독립적으로 재구현**했다
(코드 이식이 아니라 "이 엔드포인트가 이런 요청/응답을 주고받는다"는 사실만
확인한 것. NOTICE 참조):

**로그인** (`xgen_seepage/connector/auth.py`, `hash.py`):
```
POST /api/auth/login   {"email":.., "password": sha256_hex(평문), "token": null}
                        → {"success":bool, "access_token", "refresh_token", "user_id", "username"}
POST /api/auth/validate-token  {"token":.., "refresh_token":..}
                        → {"valid":bool, "user_id", "username", "roles", "permissions", "new_access_token"?}
POST /api/auth/refresh {"refresh_token":..} → {"success":bool, "access_token"}
```
비밀번호는 평문이 아니라 **SHA-256 hex 다이제스트**로 보낸다. 게이트웨이가
해시값을 그대로 비교하기 때문에, 이 규약을 안 따르면 로그인이 항상 실패한다.

**에이전트 도구 브릿지** (`xgen_seepage/connector/bridge.py`). WebSocket
`ws(s)://<서버>/api/tools/ws/connector-mcp/{user_id}`, 헤더
`Authorization: Bearer <access_token>`:
```
클라이언트→서버  {"type":"hello","catalog_id":<str>,"tools":[{"server","name","description","inputSchema"}]}
                 {"type":"ping"}                                    (20초 간격)
                 {"type":"mcp_result","request_id":..,"ok":..,"result"|"error":..}
서버→클라이언트  {"type":"ready","catalog_id":..,"tool_count":..}
                 {"type":"mcp_call","request_id":..,"server":..,"tool":..,"args":..}
```
로그인한 사용자가 XGEN 어디서(웹 UI 등) 어떤 에이전트와 채팅하든, 백엔드가 이
프로토콜을 통해 xgen-seepage가 광고한 도구(`inspect_csv`, `set_live_cell` 등)를
그 에이전트 세션에 자동으로 얹어준다. "(llm외) 에이전트들의 기능이 동작하게"
라는 요구사항이 여기서 채워진다. `tests/test_connector_bridge_integration.py`가
가짜 XGEN 서버(aiohttp)를 실제 로컬 소켓에 띄워 이 왕복 전체(로그인 → hello →
ready → mcp_call → mcp_result)를 실제 코드로 검증한다.

## 4. 독립 커넥터. 배포 모델

xgen-connector와 나란히 있는 **별개의 제품**이다. 최초 1회 `xgen-seepage login`
으로 XGEN 서버 URL + 계정으로 로그인해 토큰을 OS 키체인에 저장하고, 이후
`xgen-seepage run`이 위 WebSocket 브릿지를 상주시킨다. 이 시점부터 사용자가
XGEN 어디서 어떤 에이전트와 채팅하든, 그 에이전트가 이 프로세스에 붙은
xlwings/CSV 도구를 자동으로 쓸 수 있다. xgen-connector가 설치돼 있는지, 그
앱의 Local MCP 설정에 뭐가 들어있는지와 완전히 무관하게.

## 5. 최종 아키텍처

```
┌────────────────────────────┐
│   XGEN 서버 (게이트웨이)      │
│  /api/auth/*                │◄── xgen-seepage login (HTTPS, 1회)
│  /api/tools/ws/connector-mcp│◄── xgen-seepage run   (WSS, 상주)
└──────────────┬──────────────┘
               │ 에이전트 세션이 어디서 열리든(웹 UI 등)
               │ 이 브릿지로 도구가 자동 주입됨
┌──────────────▼──────────────┐
│  xgen-seepage (독립 프로세스) │   ← 이 프로젝트. xgen-connector 불필요
│  connector/auth.py  로그인    │
│  connector/bridge.py 도구브릿지│
│  connector/config.py 로컬설정 │
│    + OS 키체인(토큰)          │
└───────┬─────────────┬────────┘
         │             │
 ┌───────▼──┐    ┌─────▼──────────┐
 │ live_adapter│  │  csv_adapter    │
 │  (xlwings)  │  │ (csv+chardet)   │
 └───────┬─────┘  └─────┬──────────┘
         │ COM/AppleScript│ 파일 IO
 ┌───────▼──────┐  ┌──────▼─────────┐
 │ 실행 중인 Excel│  │  .csv 파일      │
 │ (열린 통합문서)│  │ (open_in_excel │
 │              │◄─┤  로 live로 승격) │
 └──────────────┘  └────────────────┘
```

인터넷이 필요한 구간은 "XGEN 서버 자체가 어디 있느냐"뿐이다. 사내망/폐쇄망에
있으면 이 커넥터의 모든 화살표가 사내망 안에서만 오간다. `xgen-seepage-mcp`
(stdio MCP 서버, `mcp_server.py`)는 이 표준 경로의 **대안**으로 남겨뒀다.
이미 xgen-connector나 Claude Desktop을 쓰고 있어서 그쪽 Local MCP 설정에 그냥
얹고 싶은 경우를 위한 것뿐, 기본 경로가 아니다.

`live_adapter`(xlwings, Excel 필요) 옆에 `libreoffice_adapter`(UNO, Excel 불필요)
가 세 번째 백엔드로 나란히 있다. 둘 다 같은 `base.py` dataclass(`CellContent`/
`SheetSchema`/`MergeInfo`)를 쓰고 `tools.py`에 별도 tool 군으로 등록돼 있어
에이전트 입장에선 어느 쪽을 부를지만 다르다. 자세한 내용은 §9.

## 6. 파일별 책임

| 파일 | 책임 |
|---|---|
| `base.py` | 공통 dataclass/예외(document-adapter와 명명 규칙 통일) |
| `_cellfmt.py` | 숫자/텍스트 판정 휴리스틱(document-adapter 이식) |
| `live_adapter.py` | xlwings로 **열려 있는** 통합문서 실시간 읽기/쓰기(Excel 필요) |
| `libreoffice_adapter.py` | UNO로 xlsx를 LibreOffice에서 직접 열어 같은 셀 IO 제공(Excel 불필요, §9) |
| `_uno_worker.py` | LibreOffice 번들 파이썬에서 도는 UNO 상주 워커(위 어댑터의 서브프로세스) |
| `csv_adapter.py` | CSV 인코딩/구분자 자동 감지 + 읽기/쓰기 + Excel로 열기 |
| `tools.py` | MCP/Anthropic tool-use 스키마 + dispatcher(document-adapter 관례) |
| `connector/hash.py` | 로그인 비밀번호 SHA-256 해싱 |
| `connector/http_client.py` | XGEN 게이트웨이용 비동기 HTTP 클라이언트 |
| `connector/auth.py` | login/validate/refresh/logout |
| `connector/config.py` | 로컬 설정 파일 + OS 키체인 토큰 저장(장애시 우아한 실패. §7) |
| `connector/bridge.py` | `/api/tools/ws/connector-mcp/{user_id}` WS 브릿지 |
| `connector/app.py` | CLI(`xgen-seepage login/logout/status/run`) |
| `mcp_server.py` | stdio MCP 서버(`xgen-seepage-mcp`, 대안 경로) |

## 7. 패키징. 폐쇄망용 단일 설치파일

`pip install`은 폐쇄망에서 원천적으로 불가능하다(조사 결과 이 회사의 기존
폐쇄망 배포 문서에 "pip/npm/cargo: 내부망에서 빌드 자체가 불가 → 반드시 개방망에서
빌드 끝내고 이미지로 반입" 이라고 명시돼 있다. 내부 레지스트리는 Docker 이미지
전용이라 PyPI 미러 같은 게 없다). 그래서 **개발자가 인터넷 있는 머신에서
PyInstaller로 미리 얼려서 단일 실행파일로 만들고, 그 실행파일만 반입**하는
모델을 쓴다. xgen-connector 자체의 electron-builder 빌드도 "빌드는 인터넷 있는
CI에서, 최종 사용자는 완성된 설치파일만" 이라는 같은 모델이라 이 방식이 이
회사의 실제 관행과 일치한다.

```powershell
pip install -e ".[dev]"
python -m PyInstaller --onefile --name xgen-seepage-connector `
  --collect-all keyring packaging/run_connector.py
```

`packaging/xgen-seepage-connector.spec`에 정확한 빌드 레시피가 있다(재현 가능).
**실제로 얼려서 검증**했다(2026-08-13, Windows 11). `--help`/`status` 정상
동작, 그리고 `packaging/smoke_test_frozen_exe.py`로 **얼린 exe가 진짜 로컬
소켓으로 가짜 XGEN 서버에 로그인하고 WS 브릿지로 hello를 보내는 것까지** 확인했다.

**실제로 찾은 버그. OS 키체인 세션 의존성**: 이 검증 과정에서 실제 버그를
하나 발견해 고쳤다. `keyring`의 Windows 백엔드는 `set_password` 내부에서
`CredRead`를 먼저 호출하는데, **대화형 로그온 세션이 아닌 컨텍스트**(이
저장소를 만든 자동화 실행 환경이 그랬다)에서는 `WinError 1312`("지정된 로그온
세션이 없습니다")를 `keyring.errors`로 감싸지 않은 raw 예외로 던져서 CLI가
스택트레이스와 함께 죽었다. `connector/config.py`에 `KeyringUnavailableError`로
감싸는 처리를 추가해 지금은 로그인 자체(HTTP)는 성공하고, 토큰 저장만 실패하면
명확한 한글 안내 메시지로 끝난다. **다만 이건 서비스 계정/원격 자동화처럼
비대화형 세션에서 도는 경우의 얘기다. 일반 사용자가 데스크톱에 설치파일을
더블클릭해서 쓰는 정상적인 경우엔 대화형 로그온 세션이 있으므로 이 문제가
안 생겨야 한다. 그 "정상 경로"는 실제 Windows 데스크톱 세션에서 별도 검증이
필요하다**. 이 저장소를 만든 환경 자체가 비대화형이라 직접 확인 못 했다.

**실제로 찾은 버그. 얼린 exe에 `_uno_worker.py`가 안 들어가 있었다**:
LibreOffice 백엔드(§9)를 붙인 뒤 실제로 얼려서 검증하기 전까지는 몰랐던
문제다. `libreoffice_adapter.py`는 `_uno_worker.py`를 import가 아니라
`Path(__file__).with_name(...)`으로 파일 경로째 찾아 서브프로세스로 띄우는데,
PyInstaller의 정적 분석은 실제로 import되는 모듈만 번들에 넣는다 - 파일
경로로만 참조되는 `.py`는 존재 자체를 모른다. 실제로 얼려서 확인해보니
그 경로가 번들 안에 없어 `worker exists = False`였다. `packaging/
xgen-seepage-connector.spec`의 `datas`에 그 파일을 명시적으로 추가해
고쳤고, 고친 뒤 얼린 exe로 진짜 LibreOffice를 열어 셀 쓰기→읽기→저장까지
왕복해 디스크에 반영된 것까지 확인했다. **폐쇄망 배포를 위해 만든 기능이
정작 폐쇄망 배포 산출물(얼린 exe) 안에서는 동작하지 않았던 것**이라 이
검증이 없었으면 실사용 시점에야 발견됐을 문제다.

**실제로 찾은 버그. `pathex=[]`라서 얼린 exe가 자기 자신을 실행조차 못 했다**
(2026-08-17): 태스크팬(§10) 패키징을 검증하려고 실제 `xgen-seepage-connector.exe`를
`packaging/` 디렉터리에서 빌드한 뒤 처음으로 **얼린 exe 자체의 CLI를 직접**
실행해봤다(`--help`) - `ModuleNotFoundError: No module named
'xgen_seepage.connector'`로 즉시 죽었다. 원인은 스펙 파일의
`pathex=[]`: PyInstaller의 정적 분석은 빌드 시점의 파이썬이 보는
`sys.path`(빌드를 실행한 디렉터리가 암묵적으로 포함됨)에 의존하는데,
`packaging/`에서 빌드하면 그 한 칸 위에 있는 `xgen_seepage/` 패키지 자체가
안 보인다. `_uno_worker.py` 문제처럼 **빌드는 성공하고 에러 없이 exe가
나오지만, 실행하면 그 자리에서 죽는** 종류라 빌드 로그만 봐서는 절대 못
잡는다. `SPECPATH`(스펙 파일이 있는 디렉터리를 가리키는 PyInstaller
내장 변수) 기준 절대경로로 `pathex`를 고정해 빌드를 어디서 실행하든
항상 프로젝트 루트를 찾도록 고쳤다.

**이 버그를 여태 놓친 이유도 정직하게 남긴다**: `_uno_worker.py` 버그를
고칠 때도, 이번 태스크팬 검증 때도, 검증은 항상 **자체 `pathex`를 따로
지정한 별도의 "probe" exe**로 했지 실제 `xgen-seepage-connector.exe`
자신의 CLI를 직접 실행한 적이 한 번도 없었다(probe는 정확한 pathex를 매번
새로 지정했으니 이 버그 자체가 안 생겼다). 그래서 이 문제가 처음
만들어진 시점(§7의 2026-08-13 "실제로 얼려서 검증. --help/status 정상
동작"이라는 기록)부터 계속 있었을 가능성이 있는데, **그 검증이 실제로
무슨 방식이었는지는 이 세션에서 재확인할 방법이 없다** - 다만 오늘
이후로는 `pathex`가 빌드 실행 위치에 안 흔들리므로 이 특정 문제는 재발
구조 자체가 없어졌다.

**위 두 버그를 고친 뒤 실제로 얼린 exe로 태스크팬 전체 왕복까지 확인했다**:
`--help`/`status`/`login`(HTTP 200, OS 키체인 저장만 위 §7 항목대로
비대화형 세션 제약으로 실패)이 정상 동작하는 것에 더해, 별도 probe exe로
(같은 spec 설정 재사용) 실제 `TaskpaneServer`를 얼린 프로세스 안에서
띄워 `/health`(200), 정적 태스크팬 디렉터리 서빙(200, 실제 HTML 응답),
`/chat/stream`(실제 dev-xgen.x2bee.com에서 SSE 6줄+ 실시간 수신)까지
전부 통과 확인했다 - uvicorn의 런타임 문자열 디스패치(`loop=auto` 등)가
걱정했던 대로 PyInstaller 정적 분석을 피해가는 문제는 이번엔 `collect_all`
+ 명시적 hidden-imports로 미리 방어해둔 덕에 실제로 재현되지 않았다.

**NSIS 설치파일로 감싸는 것(전형적인 "설치 프로그램" UX)은 아직 안 했다.**
PyInstaller onefile exe는 "실행하면 도는 단일 파일"이지 클릭 설치·시작프로그램
등록·제거 항목 같은 설치 마법사 경험은 없다. `xgen-connector`가 이미 쓰는
electron-builder NSIS 설정(`build/installer.nsh`)이 그 UX의 참고 자료가 될 수
있다(전용 설치 마법사는 Python 진영에서 보통 NSIS를 그대로, 또는 Inno Setup을
직접 쓴다. 다음 로드맵).

## 8. 실제 XGEN 서버로 끝까지 검증 (2026-08-14)

가짜 서버 상대 테스트만으로는 "진짜 되는지" 증명이 안 된다는 지적을 받고,
**실제 dev-xgen.x2bee.com 서버 + 실제 계정 + 실제 에이전트**로 전체 경로를
끝까지 확인했다. 순서대로:

1. **로그인**: 실제 계정으로 `POST /api/auth/login` 성공. `connector/auth.py`
   `hash.py`의 SHA-256 규약이 실제 게이트웨이와 정확히 일치함을 확인.
2. **WS 에이전트-도구 브릿지 첫 시도에서 진짜 버그 발견**: 연결·인증은 되는데
   `hello` 전송 직후 close 프레임 없이 연결이 끊겼다. 실제 백엔드 소스
   (GitLab `xgen-workflow` `controller/tools/connectorMcpWebSocket.py`)를
   읽어 프로토콜 자체는 맞다는 걸 확인한 뒤, 원인을 좁혀나가 **`websockets`
   라이브러리가 기본으로 협상하는 permessage-deflate 압축 확장**이 문제임을
   알아냈다(`compression=None`으로 끄니 즉시 `ready` 응답 수신). §3의 하이브리드
   프로토콜 문서에 있던 그대로였고, 문제는 클라이언트 라이브러리 기본값이었다.
3. **실제 도구 카탈로그 14개 등록 확인**: 수정 후 `server_tool_count: 14`로
   실제 서버에 정상 등록됨을 확인.
4. **실제 워크플로우로 실제 에이전트 실행**: `agents/harness` 노드 하나로 된
   최소 테스트 워크플로우를 실제로 `POST /api/agentflow/save`(정확한 요청
   스키마는 `xgen-frontend` `packages/api-client/src/agentflow.ts`에서 확인.
   `content` 래핑이 필요함)로 만들어 실제 `POST /api/agentflow/execute/
   based-id/stream`으로 실행했다. 첫 시도는 dev 서버의 vLLM 모델이 안 떠 있어
   (`503 No model loaded`) 실패했지만, 로그에 **`[Harness] Connector MCP
   도구 14개 자동 주입`**이 이미 찍혀 있어 도구 인젝션 자체는 확인됐다.
   Provider를 Anthropic(Claude Haiku 4.5)으로 바꿔 재실행하자 에이전트가
   `D:\datasets\assort\products.csv`(실제 로컬 CSV 파일)를 물었을 때 **실제로
   도구를 호출해 정확한 답을 냈다**: "81행(헤더 포함), 8열, 헤더
   product_code/product_name/.../cumulative_amount, UTF-8, 쉼표 구분". 전부
   `csv_adapter`로 직접 확인한 실제 값과 정확히 일치했다.

**즉 로그인 → 도구 등록(WS 브릿지) → 에이전트가 도구 발견 → 실제 호출 → 정확한
결과 반환까지, "(llm외) 에이전트들의 기능이 동작하게"라는 원래 요구사항의
핵심 경로가 실제 서버·실제 에이전트로 완전히 검증됐다.** 테스트에 쓴
`xgen_seepage_verify` 워크플로우는 검증 후 서버에서 삭제했다.

`bridge.py`의 `compression=None`은 `tests/test_connector_tls.py::
test_connect_always_disables_compression`로 회귀 방지 고정.

## 9. LibreOffice(UNO) 백엔드 - Excel 없는 환경 (2026-08-14)

`live_adapter`(xlwings)는 로컬에 Excel이 설치돼 있어야 동작한다. 폐쇄망
배포 대상 중엔 Excel 라이선스가 없는 환경도 흔하므로, **Excel이 없어도
같은 셀 IO 경험**을 제공하는 두 번째 백엔드로 `libreoffice_adapter.py`를
추가했다. LibreOffice의 UNO(프로세스 간 자동화 API)로 xlsx를 직접 열어
스키마 조회, 셀 읽기/쓰기, 범위 읽기/쓰기, 행 추가, 병합 셀 anchor 리다이렉트,
저장까지 `live_adapter`와 동일한 시맨틱으로 지원한다.

**아키텍처**: `uno` 파이썬 모듈은 LibreOffice가 번들한 파이썬(`program\python.exe`)
전용 네이티브 확장이라 이 패키지가 도는 일반 venv/PyInstaller exe 파이썬에서는
import가 안 된다. 그래서 `_uno_worker.py --serve`를 그 번들 파이썬으로 **상주
서브프로세스**로 띄우고, `libreoffice_adapter.py`가 stdin/stdout JSON 라인으로
명령/결과를 주고받는다. UNO 브릿지 연결은 워커 프로세스 생애주기 동안 한 번만
맺어 재사용하고, 연결 재시도·타임아웃 후 워커 재기동·병합 범위 판정 같은
견고성 처리를 워커 내부에 캡슐화해 상위 어댑터/tool 계층은 신경 쓸 필요가
없게 했다.

**범위**: 지금은 셀 읽기/쓰기/범위/행추가/저장까지만이다. 차트 생성 등은
같은 UNO 자동화 인터페이스로 자연스럽게 확장 가능하지만 이번 작업 범위 밖
(2026-08-14 사용자 확인 - "차트 만드는 것도 되겠고" 질문에 대해, 셀 IO가
핵심이고 차트는 다음 단계로).

## 10. Excel 태스크팬 채팅 (진행 중, 2026-08-17)

지금까지는 "채팅은 XGEN 어디서든(웹 UI 등), 편집은 로컬 Excel에 자동
반영"이었다. 사용자가 한 단계 더 나아가 **Excel 안에 실제 채팅 패널을
띄우는 것**을 요청했다 - §1-§2에서 설치 마찰 때문에 일부러 뺐던 Office.js
태스크팬 애드인 경로를 다시 꺼내는 것이다.

**API 재조사**: XGEN에 `/api/ai-chat/stream`이라는 가벼운 채팅 엔드포인트가
실제로 있지만, 캔버스 워크플로우 빌더와 크롬 익스텐션 전용으로 고정된
도구 세트만 쓴다 - connector-mcp 도구를 붙일 수 없어 이 경로는 못 쓴다.
우리 도구가 에이전트에 연결되는 유일한 경로는 여전히 §8에서 검증한
`POST /api/agentflow/execute/based-id/stream`(저장된 `agents/harness` 노드
워크플로우 실행)뿐이다. 그래서 태스크팬의 "채팅"은 저장된 workflow_id에
매번 `input_data`로 메시지를 보내는 것으로 구현한다.

**아키텍처**: 태스크팬은 셀을 직접 안 만지고 자체 로그인도 안 한다. 이미
도는 `xgen-seepage run` 프로세스가 로컬 HTTPS 서버(Starlette+uvicorn, 기존
WS 브릿지와 같은 이벤트 루프)를 하나 더 띄워서 (a) 태스크팬 정적 자산을
서빙하고 (b) `/chat/stream`으로 XGEN의 실행 스트림을 프록시한다. 에이전트의
tool call은 그대로 기존 WS 브릿지(`bridge.py` → `tools.call_tool()`)를
탄다 - 새로 만드는 건 UI 셸과 그 앞단 프록시뿐, 셀 IO 경로는 손대지 않는다.

**Phase 1 실측 완료**: 로컬 HTTPS 서버(`taskpane_server.py`) + 자체서명
인증서 생성(`certs.py`, `cryptography`로 순수 파이썬 생성 - `mcp`의 전이
의존성으로 이미 `.venv`에 있던 걸 확인하고 직접 의존성으로 승격)을
`cmd_run`에 연결해 기존 WS 브릿지와 `asyncio.gather`로 같이 띄웠다. 실제
dev-xgen.x2bee.com에 로그인해 브릿지가 `catalog_synced=True,
server_tool_count=24`로 실서버에 붙어있는 상태에서, 동시에 로컬
`https://127.0.0.1:51837/health`가 우리가 만든 인증서로 정상 응답하는
것까지 확인했다(겸사겸사 LibreOffice 도구 10개를 더한 뒤의 24개 카탈로그가
실서버에 정상 등록되는 것도 재확인됨).

**실측으로 찾은 진짜 문제. `certutil -addstore Root`가 이 세션에서 실패한다**:
자체서명 인증서를 Windows 신뢰 저장소에 조용히 등록하는 `certutil -user
-addstore Root`가 `0x80070032 ERROR_NOT_SUPPORTED`로 실패했다. 같은
인증서로 같은 세션에서 `-addstore CA`/`-addstore My`는 성공해서(서명
검증도 통과) 파일 포맷 문제가 아님을 확인했다 - Root 저장소 추가만 콕
집어 막힌다. OS 키체인(§7의 `KeyringUnavailableError`)이 대화형 로그온
세션이 아닌 컨텍스트에서 막히는 것과 같은 부류의 제약으로 보인다. 코드는
이미 실패를 삼키고 False를 반환하도록 방어돼 있다(`certs.py::
_trust_windows`) - 인증서 자체는 유효하니 최초 1회 브라우저/Office에서
수동으로 신뢰하는 경로는 열려 있다. **다만 일반 사용자의 정상적인
대화형 데스크톱 세션에서 이 호출이 실제로 성공하는지는 OS 키체인과
마찬가지로 별도 검증이 필요하다** - 이 저장소를 만든 환경 자체가
비대화형이라 직접 확인 못 했다.

**Phase 2 실측: 프록시 메커니즘은 증명됐다, 실행 자체는 지금 dev 서버에서
막혀 있다**: `agentflow_client.py`(하네스 워크플로우 CRUD + 실행 스트림)와
`taskpane_server.py`의 `/chat/stream`(첫 청크를 미리 당겨서 XGEN 쪽 에러를
제대로 된 HTTP 상태 코드로 바꾸고, 그 뒤로는 SSE 원시 바이트를 그대로
흘려보냄)을 만들어 실제 dev-xgen.x2bee.com으로 검증했다. **증명된 것**:
로컬 `https://127.0.0.1:PORT/chat/stream`에 POST를 보내면 실제 XGEN
서버까지 갔다가 실제 SSE 응답이 바이트 단위로 그대로 돌아온다(200, 스트림
정상 수신, 왜곡 없음) - 프록시 계층 자체는 완전히 동작한다.

**실행이 계속 `ERROR401`로 실패한 원인을 서버 로그로 끝까지 추적해 확정했다**:
클라이언트 쪽 재현(하네스 전용 API로 생성 / §8과 완전히 같은 레거시
`/api/agentflow/save` + `provider: anthropic` 명시 조합)만으로는 원인을
알 수 없어서, SSH로 실제 dev-177 클러스터의 `xgen-workflow` 파드 로그를
직접 열어봤다. 두 가지가 겹쳐 있었다:

1. **새로 만든 워크플로우는 만든 직후 자기 자신도 못 찾는다.** 하네스
   전용 API(`POST /api/agentflow/harness/workflows`)로 만들면 `{"created":
   true}`가 돌아오는데, 그 직후 같은 API의 GET/LIST/DELETE 어디로도 안
   보인다(레거시 `/api/agentflow/save`로 만든 것도 마찬가지). 파드 로그에서
   같은 부류의 실패를 실시간으로 목격했다: `apscheduler`가 정기 실행하는
   `harness_test`라는 기존 스케줄 작업조차 `ValueError: 워크플로우를 찾을
   수 없습니다`로 죽고 있었다(`execution_core.py:389`) - 즉 이건 내 계정·내
   테스트에 국한된 문제가 아니라 dev-177 환경 전반에 걸친 실제 데이터
   일관성 버그다. 원인까지는 못 좁혔지만(서버 코드 수정 권한 밖) 재현은
   명확하다.
2. **그 예외를 사용자에게 보여주려는 코드 자체가 또 죽는다.** `workflow를
   찾을 수 없습니다` 예외를 잡은 뒤 `error_message_replacer(...)`를 호출해
   에러 메시지를 만드는데(`execution_core.py:1040`), 로그에 그대로 찍혀
   있었다: `TypeError: 'NoneType' object is not callable` - 이 함수 자체가
   `None`으로 바인딩돼 있다. 그래서 진짜 원인("워크플로우 없음")이 감춰지고
   내 클라이언트에는 의미 없는 제네릭 `ERROR401`만 도달했다.

**실측으로 원인을 분리하려고 만든 새 워크플로우를 버리고, 대신 실제로
이미 존재가 확인된 워크플로우(`shinhan_blue_agent_v1`)로 그대로 다시
시도하자 - 완전히 정상 동작했다.** `event: log`/`node_status`/
`execution_io` 전체가 실시간으로 로컬 프록시를 거쳐 정확히 도착했고, 워크플로우
자체는 끝까지(`"Stream finished"`) 돌았다. 유일하게 남은 실패는 그
워크플로우의 harness 노드가 기본으로 쓰는 `provider=vllm, model=Qwen3.6-27B`
에서 `Server error 503: No model loaded`였다 - **§8에서 이미 3일 전에
확인했던, 이 dev 서버의 그 로컬 모델 미기동 문제와 완전히 같은 것**이지
새로운 문제가 아니다.

**이 실측이 설계 자체를 바로잡았다(사용자 정정, 2026-08-17)**: 애초에
"채팅용 워크플로우가 없으면 xgen-seepage가 자동 생성"하는 Phase 3을
계획했었는데, 틀린 방향이었다. xgen-connector가 하듯 **연결**이어야 한다 -
사용자가 XGEN에 이미 갖고 있는 워크플로우 중 아무거나 골라 태스크팬을
거기 연결하는 것이지, xgen-seepage가 전용 워크플로우를 서버에 몰래
만들어 소유하는 게 아니다. 이 방향이 위 버그도 자동으로 비켜간다 - 이미
존재가 검증된 워크플로우만 다루므로 "방금 만든 게 안 보이는" 문제 자체가
발생할 여지가 없다. `agentflow_client.py::list_workflows()`(`GET
/api/agentflow/list`)와 새 CLI `xgen-seepage chat-workflow list`/`set`으로
구현했고, 실제로 `shinhan_blue_agent_v1`을 이 명령으로 연결하는 것까지
실측했다. `config.py`의 이제는 안 쓰는 `chat_provider`/`chat_model`
필드(자동생성 템플릿용이었다)는 제거했다.

**Phase 5(패키징)는 완료됐다** - §7에 기록한 `pathex` 버그 발견/수정과
얼린 exe로의 태스크팬 전체 왕복 검증이 그 내용이다. **남은 단계는 Phase 4
(실제 Excel에 태스크팬)뿐이다 - Excel이 없는 이 머신에서는 구조적으로
진행 불가**, §11 하단의 Excel COM 미검증 항목과 같은 성격. 게다가
Office.js 자체가 인터넷 전혀 없이 로드되는지도 MS 공식 문서로도
불확실해서(자체 호스팅 비권장, 실제 시도자는 인터넷 끄니 안 됐다고 보고)
실제 Excel 머신에서 함께 확인해야 한다. "워크플로우 자동생성"이던 옛
Phase 3은 폐기 - 연결 메커니즘으로 이미 대체됐다.

## 11. 알려진 제약 / 로드맵

- **폐쇄망 반입 경로 미정. 실제 조사 결과**: "제주은행" 폐쇄망 사례를 조사해보니,
  이 회사의 기존 폐쇄망 반입은 **전부 Docker 이미지**로 이뤄진다(USB로 물리 반입
  → `docker load` → 내부 레지스트리 → k3s). **데스크톱 설치파일(.exe)이 이
  경로를 탄 전례가 전혀 없다.** xgen-seepage는 "일단 특정 고객사 아님, 범용
  기능으로 설계"(2026-08-13 사용자 확인)라 지금은 이 문제를 풀지 않지만, 실제
  폐쇄망에 반입할 때는 **desktop .exe 반입이라는 새 반입 유형** 자체를 그
  네트워크 운영자와 협의해야 한다. USB로 파일 자체는 옮길 수 있어 보이지만
  (예: 다른 서드파티 이미지도 USB로 들어간 전례가 있다), 보안심의 절차가
  Docker 이미지 기준으로 짜여 있어 그대로 통과될지는 미검증이다.
- **OS 키체인 세션 의존성**: §7 참조. 서비스 계정/원격 실행 컨텍스트에서
  `xgen-seepage login`을 돌리면 토큰 저장이 실패할 수 있다(에러 메시지는
  명확하게 나온다). 일반 데스크톱 세션에서의 정상 동작은 별도 검증 필요.
- **성능**: `live_adapter`의 병합 셀 스캔은 COM 호출이 셀당 1회다. 병합이 아예
  없는 시트는 `UsedRange.MergeCells == False` 한 번의 호출로 빠르게 판정하지만,
  병합이 있는 큰 시트는 `_MAX_MERGE_SCAN_CELLS`(20,000) 상한에 걸려
  `truncated=True`를 반환한다. `read_range`/`write_range`(벌크 API)를 넓은
  범위에 쓰도록 tool description에 명시했다.
- **닫힌 xlsx 파일 편집**: 이 프로젝트의 범위가 아니다. 필요하면 `document-adapter`
  를 별도로 쓰면 된다(같은 dataclass 명명 규칙이라 에이전트 입장에서 다루기 쉽다).
- **macOS 미검증**: xlwings의 AppleScript 경로는 API상 동일하게 동작해야 하지만
  실제 검증하지 못했다.
- **Windows COM, 실제 Excel 붙여서 하는 E2E는 여전히 미검증**: 이 저장소를
  만든 SERVER_JS 머신에는 Microsoft Excel이 설치돼 있지 않아(§9) `live_adapter`가
  실제로 열려 있는 Excel 통합문서의 셀을 읽고/쓰는 COM 경로 자체는 검증하지
  못했다. §9의 LibreOffice 백엔드가 "Excel 없는 환경에서 셀 IO"라는 요구사항
  자체는 채우므로 우선순위는 낮다. 필요하면 Excel이 실제 설치된 Windows
  머신에서 `get_live_schema`/`set_live_cell`/`read_live_range`를 열린 파일로
  직접 스모크 테스트.
- **NSIS/설치 마법사 UX**: §7 참조, 아직 onefile exe 단계.
