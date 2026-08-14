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

## 10. 알려진 제약 / 로드맵

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
