"""xgen-seepage. 폐쇄망에서도 설치되는 독립형 XGEN Excel/CSV 커넥터.

세 실행 경로:
  - `live_adapter`: 지금 로컬 Excel에서 열려 있는 통합문서를 xlwings(COM/AppleScript)로
    실시간 편집한다(수식 포함). 인터넷·웹서버·매니페스트가 필요 없다.
  - `libreoffice_adapter`: Excel이 없는 환경(폐쇄망 등)에서 xlsx를 LibreOffice(UNO)로
    직접 열어 같은 셀 IO 경험을 제공한다.
  - `csv_adapter`: CSV 파일을 인코딩/구분자 자동 감지로 읽고 편집한다. 라이브
    Excel로 올리려면 `csv_adapter.open_in_excel`로 넘어간다.

세 어댑터는 `tools.py`에서 MCP tool로 노출되고, `connector/`가 XGEN 서버에
직접 로그인해 에이전트-도구 WebSocket 브릿지로 광고한다(`xgen-seepage`
CLI). 별도 데스크톱 커넥터 클라이언트를 설치할 필요가 없다.
"""

__version__ = "0.21.0"
