# Changelog

## 0.2.0 — 2026-08-13

방향 전환: xgen-connector의 Local MCP에 얹는 방식에서, **독립 실행형
커넥터**로 재설계(사용자 정정 — 폐쇄망 단일 설치 프로그램이어야 하고,
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
  검증하는 수동 스모크 스크립트. **실제로 얼려서 검증** — 그 과정에서 OS
  키체인이 비대화형 세션에서 raw 예외로 CLI를 죽이는 실제 버그를 찾아
  `KeyringUnavailableError`로 감싸 고침.
- `ARCHITECTURE.md`: 독립 프로토콜 재구현 근거, 패키징/검증 기록, 폐쇄망
  반입 경로 조사 결과(제주은행 사례 — Docker 이미지 전용 반입, 데스크톱
  설치파일 전례 없음. 미해결로 로드맵에 기록) 추가.

## 0.1.0 — 2026-08-13

Initial scaffold.

- `live_adapter`: xlwings 기반, 로컬에서 열려 있는 Excel 통합문서를 실시간으로
  읽고 쓴다(값·수식 모두, 병합 셀 인지, 범위 벌크 읽기/쓰기).
- `csv_adapter`: 인코딩(BOM/chardet)·구분자 자동 감지 CSV 읽기/쓰기, 셀 단위
  편집, Excel로 열어 라이브 모드로 넘기는 `open_in_excel`.
- `tools.py` / `mcp_server.py`: 14개 도구를 MCP stdio 서버로 노출
  (`xgen-seepage-mcp`). XGEN Connector의 "로컬 MCP" 브릿지에 그대로 등록 가능.
- 리서치 기반 아키텍처 문서 `ARCHITECTURE.md`.
