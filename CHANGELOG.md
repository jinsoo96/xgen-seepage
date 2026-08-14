# Changelog

## 0.3.0 (2026-08-14)

**LibreOffice(UNO) 백엔드 추가. Microsoft Excel이 없는 환경(폐쇄망 등)에서도
같은 셀 IO 경험을 제공.** 이 저장소를 만든 머신엔 Excel이 없다(3가지 방법으로
확인). winget으로 Office 설치를 시도했으나 이 세션의 비관리자 권한 제약으로
막혔고, 대신 이미 설치돼 있던 LibreOffice로 실제 동작하는 대안 백엔드를
새로 만들었다(자세한 아키텍처/버그 4가지는 `ARCHITECTURE.md` §9).

- `libreoffice_adapter.py` / `_uno_worker.py` 신규: UNO로 xlsx를 LibreOffice에서
  직접 열어 스키마 조회/셀 읽기·쓰기/범위 읽기·쓰기/행 추가/저장을 제공.
  `uno` 모듈이 LibreOffice 번들 파이썬 전용이라 상주 서브프로세스(stdin/stdout
  JSON 프로토콜)로 격리했다.
- `tools.py`에 `open_libreoffice_document`/`get_libreoffice_schema`/
  `set_libreoffice_cell`/`read_libreoffice_range`/`write_libreoffice_range`/
  `append_libreoffice_row`/`save_libreoffice_document` 등 10개 도구를 별도
  tool 군으로 등록(총 24개 도구).
- `tests/test_libreoffice_adapter.py` 신규: 실제 LibreOffice 애플리케이션으로
  왕복(연결→쓰기→읽기→저장→디스크 검증)까지 도는 13개 테스트. `live_adapter`의
  xlwings 테스트(Excel 없어 우아한 실패만 검증)와 달리 이건 실기 E2E다.
  56개 전체 테스트 통과, ruff/mypy 클린.

## 0.2.2 (2026-08-14)

**실제 XGEN dev 서버(dev-xgen.x2bee.com) 상대 완전한 엔드투엔드 검증.**
가짜 서버 테스트만으로는 부족하다는 지적에 따라 실제 계정으로 로그인부터
실제 에이전트의 도구 호출까지 전부 실측했다(자세한 과정은 `ARCHITECTURE.md`
§8).

- **실제 버그 수정**: `bridge.py`의 `websockets.connect()`가 기본값인
  permessage-deflate 압축 확장을 협상하면, 실제 서버 환경에서 `hello`
  전송 직후 close 프레임 없이 연결이 끊겼다. `compression=None`으로 고정.
  회귀 테스트 `tests/test_connector_tls.py::
  test_connect_always_disables_compression` 추가.
- 실제 서버에 실제 도구 카탈로그 14개 등록 확인(`server_tool_count: 14`).
- 실제 워크플로우(agents/harness 노드)를 실제로 저장·실행해, 실제 에이전트
  (Claude Haiku 4.5)가 `mcp_xgen-seepage_inspect_csv`를 실제로 호출해
  로컬 CSV 파일(`D:\datasets\assort\products.csv`)의 정확한 구조(81행,
  8열, 헤더 전체, 인코딩, 구분자)를 답하는 것까지 확인. 테스트에 쓴 워크플로우는
  검증 후 서버에서 삭제.
- 44개 단위/통합 테스트 통과(로컬), 실제 서버 상대 수동 검증 별도 완료.

## 0.2.1 (2026-08-13)

- 사용자 요청으로 xgen-connector의 연결 로직을 더 적극적으로 참고: 사설/
  자체서명 CA를 쓰는 폐쇄망 XGEN 서버 접속을 위한 `allow_private_certificate`
  를 실제로 구현했다(전에는 설정 필드만 있고 어디에도 안 이어져 있었다).
  `connection_security.py` 신규(`connection-security.ts`의
  `shouldAllowPrivateCertificate`/`xgenWebSocketTlsOptions`와 같은 취지로
  포팅, NOTICE 참조). `login --allow-private-certificate` 플래그 +
  `XGEN_SEEPAGE_ALLOW_PRIVATE_CERTIFICATE` 환경변수(배포 기본값용, 역시
  xgen-connector의 `deployment-defaults.ts` 패턴 참고) 추가. wss://에만
  적용하고 ws://에는 안 넣는 로직까지 실제 테스트로 검증
  (`tests/test_connector_tls.py`).
- **실제 크래시 버그 발견 및 전면 수정**: `argparse` 도움말 문자열에 든
  em dash(U+2014)가 Windows 한글 콘솔(cp949) 출력 시 `UnicodeEncodeError`로
  CLI 자체를 죽였다(`xgen-seepage login --help` 재현). 이 프로젝트 전체
  (코드 주석·문서·커밋 예정 텍스트 포함)에서 em dash를 전부 제거했다
  (사용자 지침 `no-em-dash` 위반이기도 했음. 42/42 테스트 통과 유지 확인).

## 0.2.0 (2026-08-13)

방향 전환: xgen-connector의 Local MCP에 얹는 방식에서, **독립 실행형
커넥터**로 재설계(사용자 정정. 폐쇄망 단일 설치 프로그램이어야 하고,
document-adapter/xgen-doc2chunk/xgen-connector는 참고 자료일 뿐 런타임
의존성이면 안 됨).

- `connector/` 신규: XGEN 서버에 직접 로그인(`hash.py`/`auth.py`,
  `/api/auth/login` 등 실제 프로토콜을 xgen-connector 소스에서 확인해 독립
  재구현) + `/api/tools/ws/connector-mcp/{user_id}` WebSocket 브릿지
  (`bridge.py`, hello/ready/mcp_call/mcp_result) + 로컬 설정/OS 키체인 토큰
  저장(`config.py`) + CLI(`app.py`: `login`/`run`/`status`/`logout`).
- 새 콘솔 스크립트 `xgen-seepage`(기본 경로). 기존 `xgen-seepage-mcp`는
  xgen-connector/Claude Desktop에 얹고 싶을 때 쓰는 대안으로 유지.
- `tests/test_connector_*.py`: 해싱/설정파일/OS 키체인 왕복 단위테스트 +
  가짜 XGEN 서버(aiohttp, 실제 로컬 소켓)로 로그인+WS 브릿지 전체 왕복을
  검증하는 통합테스트. 총 36개 테스트 통과.
- `packaging/`: PyInstaller onefile 빌드 레시피(`run_connector.py`,
  `xgen-seepage-connector.spec`) + 얼린 exe를 진짜 서브프로세스로 띄워
  검증하는 수동 스모크 스크립트. **실제로 얼려서 검증**. 그 과정에서 OS
  키체인이 비대화형 세션에서 raw 예외로 CLI를 죽이는 실제 버그를 찾아
  `KeyringUnavailableError`로 감싸 고침.
- `ARCHITECTURE.md`: 독립 프로토콜 재구현 근거, 패키징/검증 기록, 폐쇄망
  반입 경로 조사 결과(제주은행 사례. Docker 이미지 전용 반입, 데스크톱
  설치파일 전례 없음. 미해결로 로드맵에 기록) 추가.

## 0.1.0 (2026-08-13)

Initial scaffold.

- `live_adapter`: xlwings 기반, 로컬에서 열려 있는 Excel 통합문서를 실시간으로
  읽고 쓴다(값·수식 모두, 병합 셀 인지, 범위 벌크 읽기/쓰기).
- `csv_adapter`: 인코딩(BOM/chardet)·구분자 자동 감지 CSV 읽기/쓰기, 셀 단위
  편집, Excel로 열어 라이브 모드로 넘기는 `open_in_excel`.
- `tools.py` / `mcp_server.py`: 14개 도구를 MCP stdio 서버로 노출
  (`xgen-seepage-mcp`). XGEN Connector의 "로컬 MCP" 브릿지에 그대로 등록 가능.
- 리서치 기반 아키텍처 문서 `ARCHITECTURE.md`.
