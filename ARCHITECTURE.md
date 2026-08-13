# xgen-seepage — 아키텍처 & 리서치 노트

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
클라우드(`api.anthropic.com`)이고, 로그인한 유료 Claude 계정이 필요하다 — **오프라인
동작이 아니다.**

### ChatGPT for Excel / 서드파티 (GPT for Excel 등)

같은 패턴(Office.js 태스크팬) + 추가로 `=GPT()` 같은 **Custom Function**으로 셀
수식처럼 호출하는 방식도 있다. 전부 클라우드 LLM 호출이 전제다.

### Office Add-ins 플랫폼 자체는 오프라인이 가능하다

두 제품이 클라우드에 묶여 있다고 해서 **플랫폼 자체**가 오프라인 불가능한 건
아니다:
- 태스크팬은 그냥 HTTPS로 호스팅되는 웹페이지다. `localhost`에 자체서명 인증서
  (`office-addin-dev-certs`)를 붙이면 Office의 내장 WebView2가 문제없이 로드한다 —
  즉 **원리적으로는** 완전 로컬 웹서버 + Office.js 태스크팬으로 오프라인 커넥터를
  만들 수 있다.
- Office Store 없이도 **네트워크 공유 폴더 카탈로그**(XML 매니페스트, JSON 통합
  매니페스트는 이 경로 미지원)로 사이드로딩할 수 있다.
- 셀 API: `range.load("values,formulas,formulasR1C1")` + `context.sync()`로 값과
  수식을 분리해서 읽고, `range.values =` / `range.formulas =`로 쓴다.

**그런데 이 경로를 그대로 택하지 않은 이유**: 인증서 신뢰, 매니페스트 사이드로딩,
SharedRuntime 설정 등 설치 마찰이 여전히 크다. "인터넷 없이도 설치"라는 목표에
비해 얻는 게 적다 — 태스크팬은 **UI 셸**일 뿐이고, 우리가 필요한 건 셀 IO다.

## 2. 대안: xlwings (채택)

`xlwings`는 Windows는 COM, macOS는 AppleScript로 **로컬에서 실행 중인 Excel
프로세스에 직접 붙는다.** 웹서버도, 인증서도, 매니페스트도 없다 — `pip install`만
하면 끝나는 순수 로컬 프로세스 간 통신이라 **인터넷 요구사항이 원천적으로 없다.**

| | Office.js 태스크팬 (Claude/ChatGPT 실제 구현) | xlwings COM/AppleScript (이 프로젝트) |
|---|---|---|
| 설치 | 매니페스트 사이드로딩 + 인증서 신뢰 | `pip install` |
| 인터넷 | 태스크팬 자체는 불필요(localhost 가능) | 불필요 |
| UI | 시트 옆 패널(진짜 "붙어 있는" 느낌) | 없음(별도 채팅 UI 필요 — 아래 §4) |
| 셀 IO | `range.values` / `range.formulas` | `Range.value` / `Range.formula`/`.formula2` |
| 플랫폼 | Windows/Mac/Web Excel | Windows/Mac 데스크톱 Excel만(Web 불가) |

값/수식 API가 거의 1:1 대응이라(`values`↔`.value`, `formulas`↔`.formula`) 편집
시맨틱을 옮기는 데 문제가 없었다. 대신 잃는 건 "시트 옆에 항상 붙어 있는 패널" UI다
— 이건 아키텍처가 아니라 UI 계층의 문제이므로 §4에서 별도로 다룬다.

## 3. PlateerLab 생태계에서 재사용한 것 (레거시 먼저 따라라)

새로 만들기 전에 회사 GitHub(`PlateerLab`)에 이미 있는 것부터 뒤졌다. 세 레포가
직접적인 근거가 됐다.

### `PlateerLab/document-adapter` — 셀 편집 시맨틱의 원본

이미 **닫힌 xlsx 파일**을 openpyxl로 열어 병합 셀 인지, 수식 캐시값/수식 문자열
구분, 숫자-처럼-보이는-문자열 자동 숫자화, MCP 서버 노출까지 다 구현해 놓은
Apache-2.0 라이브러리다(Claude Desktop/Code에서 바로 MCP로 붙는다). **이 프로젝트가
새로 만든 건 "파일이 아니라 지금 열려 있는 통합문서"라는 한 가지뿐이다** — 그
차이를 만드는 편집 시맨틱(병합 셀 anchor 리다이렉트, 숫자/문자 판정 휴리스틱)은
`document-adapter/xlsx_adapter.py`에서 그대로 이식했다(`xgen_seepage/_cellfmt.py`,
`live_adapter.py`의 `MergedCellWriteError` 처리 — 출처는 `NOTICE` 참조).
document-adapter 자체는 의존성으로 새로 추가하지 않았다: **"닫힌 파일" 편집이
필요하면 document-adapter를 그대로 쓰면 되고**, xgen-seepage가 중복 구현할
이유가 없기 때문이다(v0.2 로드맵 참고).

### `PlateerLab/xgen-doc2chunk` — CSV 인코딩 감지 전략

국내 실사용 CSV(EUC-KR/CP949로 저장된 공공·사내 데이터)를 다루면서 검증된
"BOM 우선 → chardet(신뢰도 임계값) → 후보 인코딩 순차 시도 → latin-1 최종 폴백"
순서를 그대로 이식했다(`xgen_seepage/csv_adapter.py::detect_encoding`, 출처는
`NOTICE`). 새로 설계하지 않고 이미 실전 검증된 순서를 썼다.

### `PlateerLab/xgen-connector` — 배포/통합 지점

이 프로젝트가 **가장 크게 기댄 기존 인프라**다. `xgen-connector`(Electron 데스크톱
앱)는 이미 **"로컬 MCP" 브릿지**를 갖고 있다: 사용자가 로컬 MCP 서버(stdio 또는
Streamable HTTP)를 등록하면, `mcp-manager.ts`가 그 서버에 클라이언트로 붙어 도구
카탈로그를 모으고, `mcp-bridge.ts`가 `/api/tools/ws/connector-mcp/{user_id}`
웹소켓으로 XGEN 백엔드에 알려서, 로그인한 사용자가 어떤 XGEN 에이전트
(`agent_xgen`/`agent_harness`/`agent_geny`)와 채팅하든 그 도구들이 자동 주입된다.

**즉, xgen-seepage가 "잘 만든 로컬 MCP 서버"이기만 하면, 채팅 UI·인증·에이전트
연결은 전부 xgen-connector가 이미 해결한 문제다.** 그래서 이 프로젝트는 태스크팬이나
자체 데스크톱 셸을 새로 만들지 않고, `xgen-seepage-mcp`(stdio MCP 서버) 하나에
집중했다. 연동 방법은 `examples/xgen_connector_local_mcp.md` 참조.

## 4. 최종 아키텍처

```
                     ┌─────────────────────────────┐
                     │   XGEN 에이전트 세션          │
                     │ (agent_xgen/harness/geny)    │
                     └──────────────┬────────────────┘
                                    │ WebSocket
                                    │ /api/tools/ws/connector-mcp/{user_id}
                     ┌──────────────▼────────────────┐
                     │   xgen-connector (Electron)    │  ← 이미 존재, 재사용
                     │   mcp-bridge.ts / mcp-manager  │
                     └──────────────┬────────────────┘
                                    │ stdio (로컬 프로세스)
                     ┌──────────────▼────────────────┐
                     │   xgen-seepage-mcp             │  ← 이 프로젝트
                     │   tools.py → call_tool(name,…) │
                     └───────┬─────────────┬──────────┘
                              │             │
                  ┌───────────▼──┐    ┌─────▼──────────┐
                  │ live_adapter │    │  csv_adapter    │
                  │  (xlwings)   │    │ (csv+chardet)   │
                  └───────┬──────┘    └─────┬──────────┘
                          │ COM/AppleScript  │ 파일 IO
                  ┌───────▼──────┐    ┌──────▼─────────┐
                  │ 실행 중인 Excel│    │  .csv 파일     │
                  │ (열린 통합문서)│    │ (open_in_excel │
                  │               │◄───┤  로 live로 승격 가능)│
                  └───────────────┘    └────────────────┘
```

인터넷이 필요한 구간이 하나도 없다 — XGEN 백엔드 자체가 사내망/로컬에 있다는
전제 하에, 이 커넥터의 모든 화살표는 로컬 프로세스 간 통신이다.

## 5. 파일별 책임

| 파일 | 책임 |
|---|---|
| `base.py` | 공통 dataclass/예외(document-adapter와 명명 규칙 통일) |
| `_cellfmt.py` | 숫자/텍스트 판정 휴리스틱(document-adapter 이식) |
| `live_adapter.py` | xlwings로 **열려 있는** 통합문서 실시간 읽기/쓰기 |
| `csv_adapter.py` | CSV 인코딩/구분자 자동 감지 + 읽기/쓰기 + Excel로 열기 |
| `tools.py` | MCP/Anthropic tool-use 스키마 + dispatcher(document-adapter 관례) |
| `mcp_server.py` | stdio MCP 서버(`xgen-seepage-mcp`) |

## 6. 알려진 제약 / v0.2 로드맵

- **성능**: `live_adapter`의 병합 셀 스캔은 COM 호출이 셀당 1회다. 병합이 아예
  없는 시트는 `UsedRange.MergeCells == False` 한 번의 호출로 빠르게 판정하지만,
  병합이 있는 큰 시트는 `_MAX_MERGE_SCAN_CELLS`(20,000) 상한에 걸려
  `truncated=True`를 반환한다 — 범위를 좁혀 재조회해야 한다. `read_range`/
  `write_range`는 벌크 API라 훨씬 빠르므로, 넓은 범위는 셀 단위 도구 대신 이걸
  쓰도록 tool description에 명시했다.
- **닫힌 xlsx 파일 편집**: 이 프로젝트의 범위가 아니다. `document-adapter`가 이미
  해결했으므로, 파일 모드가 필요하면 `document-adapter`를 xgen-connector의 로컬
  MCP에 **같이** 등록하면 된다(같은 dataclass 명명 규칙을 써서 에이전트가 두
  서버를 자연스럽게 같이 쓸 수 있게 했다).
- **macOS 미검증**: xlwings의 AppleScript 경로는 API상 동일하게 동작해야 하지만,
  이 리포를 만든 환경(Windows, Excel 미설치)에서는 실제 검증하지 못했다.
- **Windows COM, 실제 Excel 붙여서 하는 E2E 미검증**: 이 저장소를 만든 SERVER_JS
  머신에는 Microsoft Excel이 설치돼 있지 않다(레지스트리·`excel.exe` 둘 다 없음,
  2026-08-13 확인). `csv_adapter`와 `tools.py` 디스패처, 그리고
  `live_adapter`의 "실행 중인 Excel이 없을 때" 우아한 실패 경로(`xw.apps == []`
  → `WorkbookNotFoundError`)는 전부 실제로 돌려서 검증했다(`tests/`). 하지만
  **실제로 열린 통합문서의 셀을 읽고/쓰는 경로는 Excel이 설치된 머신에서 별도
  검증이 필요하다** — xlwings API 문서와 document-adapter의 동일 시맨틱을
  근거로 구현했지만, COM 타입 변환(날짜/병합/서식)의 실제 동작은 문서와 실기가
  미묘하게 다를 수 있다.
