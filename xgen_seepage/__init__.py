"""xgen-seepage — 오프라인 설치형 XGEN Excel/CSV 커넥터.

두 실행 경로:
  - `live_adapter`: 지금 로컬 Excel에서 열려 있는 통합문서를 xlwings(COM/AppleScript)로
    실시간 편집한다(수식 포함). 인터넷·웹서버·매니페스트가 필요 없다.
  - `csv_adapter`: CSV 파일을 인코딩/구분자 자동 감지로 읽고 편집한다. 라이브
    Excel로 올리려면 `csv_adapter.open_in_excel`로 넘어간다.

두 어댑터는 `tools.py`에서 MCP tool로 노출되며, `xgen-connector`(PlateerLab)의
"로컬 MCP" 브릿지에 등록하면 XGEN 에이전트 세션에 자동 연결된다.
"""

__version__ = "0.1.0"
