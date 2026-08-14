# 2026-08-14: LibreOffice(UNO) 백엔드 구축 - Excel 없는 환경 대응

## 배경

이전 워크로그(`2026-08-14_실서버_엔드투엔드_검증.md`)에서 로그인→WS 브릿지→
에이전트 도구 호출까지는 실제 서버로 검증했지만, `live_adapter`(xlwings/COM)가
실제 Excel 통합문서 셀을 읽고/쓰는 것 자체는 이 머신(SERVER_JS)에 Excel이
없어 미검증으로 남아 있었다. Office 설치를 시도했으나 이 세션엔 실제 admin
권한이 없어(`C:\Program Files`에 쓰기 실패로 직접 확인) 막혔다. 이 머신엔
LibreOffice가 이미 설치돼 있었으므로, Excel이 없어도 같은 셀 IO를 제공하는
대안 백엔드를 만드는 방향으로 풀었다.

## 구현

`libreoffice_adapter.py`가 LibreOffice의 UNO(프로세스 간 자동화 API)로
xlsx를 직접 열어 `live_adapter`와 동일한 시맨틱(스키마 조회, 셀 읽기/쓰기,
범위 읽기/쓰기, 행 추가, 병합 셀 anchor 리다이렉트, 저장)을 제공한다. `uno`
모듈은 LibreOffice 번들 파이썬 전용 네이티브 확장이라 일반 venv에서 import가
안 되므로, `_uno_worker.py --serve`를 그 번들 파이썬으로 상주 서브프로세스로
띄우고 stdin/stdout JSON 라인 프로토콜로 통신하는 구조로 분리했다. UNO 브릿지
연결은 워커 생애주기 동안 한 번만 맺어 재사용하고, 연결 재시도·타임아웃 시
워커 재기동·병합 범위 판정 같은 견고성 처리는 워커 내부에 캡슐화했다.

`tools.py`엔 `live_adapter`/`csv_adapter`와 동일한 패턴(dict 반환 래퍼 +
TOOL_DEFINITIONS + TOOL_HANDLERS)으로 `open_libreoffice_document`부터
`save_libreoffice_document`까지 10개 도구를 별도 tool 군으로 등록했다.
총 도구 수 14→24.

## 검증

`D:\Jeju_test\JGP-TE-GP-10-결함관리대장_v0.1_20260716.xlsx`(실제 제주은행
결함관리대장) 사본으로 실제 애플리케이션 왕복 편집을 확인했다: 마커 쓰기 →
읽기로 값 확인 → 원복 → 저장 없이 닫기, 원본 파일은 손대지 않았다. 병합 셀
anchor/non-anchor 판정, 수식 쓰기, range 읽기/쓰기, 행 추가, 미지 문서 예외,
저장 후 디스크 반영까지 같은 방식으로 실제 파일 상대로 확인했다. 기존
43개 단위/통합 테스트는 그대로 유지, ruff/mypy 클린.

## 환경 특이사항

Bash 툴로 띄운 soffice.exe는 백그라운드 프로세스가 재핑(reap)당해 재현이
불안정했다(동일 코드가 단독 실행 시엔 됨). PowerShell 툴로 전환하니 안정적으로
유지됨 - 이후 LibreOffice 관련 실행은 전부 PowerShell 툴로 통일.

## 결론

Excel이 없는 환경(폐쇄망 등)에서도 xlsx 파일에 실제로 침투해 에이전트가
셀을 읽고/쓰는 전체 경로가 실기로 증명됐다. 버전 0.3.0.

## 남은 것

`live_adapter`(xlwings/COM) 자체는 여전히 미검증(Excel이 없는 이 머신에서는
구조적으로 불가). LibreOffice 백엔드가 "Excel 없는 환경에서도 셀 IO"라는
요구사항 자체를 채우므로 우선순위는 낮아짐. 차트 생성 등은 같은 UNO 인터페이스로
확장 가능하지만 이번 범위 밖(사용자 확인, `ARCHITECTURE.md` §9 참고).
