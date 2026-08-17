# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)

**폐쇄망에서도 설치되는 독립형 XGEN Excel/CSV 커넥터.** 챗지피티나 클로드가
엑셀 안에 플러그인처럼 붙어서 편집·붙여넣기하는 것의 XGEN 버전. XGEN 서버에
직접 로그인해서 붙는 **독립 실행 프로그램**이고(다른 앱 설치 불필요), 인터넷
없이도 설치·운영 가능하다.

- **로그인 한 번**(`xgen-seepage login`) → XGEN 서버·계정에 붙는다. 여러
  XGEN(jeju/dev/prod)을 오갈 땐 로그인 시 써 본 서버 목록에서 고른다.
- **`xgen-seepage run`을 상주**시키면, 사용자가 XGEN 어디서 어떤 에이전트와
  채팅하든 그 에이전트가 사용자 로컬 Excel/CSV를 도구로 쓸 수 있게 된다.
- 지금 로컬 Excel에서 **열려 있는** 통합문서를 실시간으로 읽고 쓴다(수식 포함,
  파일 저장 불필요. 사용자가 화면에서 바로 결과를 본다). **실측 확인**: 별도
  프로세스로 도는 커넥터가 사용자가 연 통합문서를 찾아 에이전트가 셀을 실제로
  바꾼다(`ARCHITECTURE.md` §11).
- **Excel 안 채팅 패널**: `xgen-seepage panel`이 브라우저로 채팅 패널을 연다.
  패널에서 **에이전트(워크플로우)와 모델(provider)을 골라** 채팅하면 그
  에이전트가 열린 Excel을 편집한다. 이 패널은 office.js에 의존하지 않는 순수
  웹 UI라 XGEN 로그인에만 의존한다(인터넷·Microsoft 애드인 인프라 불필요).
  Office가 지원하면 Excel 리본 버튼으로도 띄울 수 있다(`taskpane/manifest.xml`).
- **Excel이 없는 환경**(폐쇄망 등)에서도 LibreOffice로 xlsx를 직접 열어 같은
  셀 읽기/쓰기를 제공한다. 파일을 미리 열어둘 필요도 없다.
- CSV는 인코딩/구분자를 자동 감지해 읽고 편집하며, 필요하면 Excel로 올려 같은
  라이브 경로로 넘어간다.

📦 자세한 아키텍처·설계 근거는 [ARCHITECTURE.md](ARCHITECTURE.md).

## 설치 & 사용

**개발/테스트용 (인터넷 있는 환경):**

```bash
pip install "xgen-seepage[live]"   # 라이브(열려 있는 통합문서) 편집 포함
xgen-seepage login                 # XGEN 서버 URL + 계정 (대화형 프롬프트)
xgen-seepage run                   # 브릿지 상주 시작 (Ctrl+C로 종료)
```

**폐쇄망용 (단일 설치파일):** `packaging/` 아래에 PyInstaller 빌드 레시피가
있다. 인터넷 있는 머신에서 미리 얼려서 `xgen-seepage-connector.exe` 하나로
만든 뒤, 그 실행파일만 폐쇄망에 반입한다.

```powershell
pip install -e ".[dev]"
python -m PyInstaller --onefile --name xgen-seepage-connector `
  --collect-all keyring packaging/run_connector.py
```

빌드 산출물(`packaging/dist/xgen-seepage-connector.exe`)은 어떤 파이썬도
설치돼 있지 않은 Windows 머신에서 그대로 실행된다. `.exe login` /
`.exe run` / `.exe status` / `.exe logout`.

**비대화형 배포**(대량 설치 시 로그인을 스크립트로): 환경변수
`XGEN_SEEPAGE_SERVER_URL` / `XGEN_SEEPAGE_EMAIL` / `XGEN_SEEPAGE_PASSWORD`로
`login`을 자동화할 수 있다.

## CLI

| 명령 | 설명 |
|---|---|
| `xgen-seepage login` | XGEN 서버에 로그인(써 본 서버 목록에서 선택), 토큰을 OS 키체인에 저장 |
| `xgen-seepage run` | 에이전트-도구 브릿지 + 채팅 패널 로컬 서버를 상주시킨다(포그라운드). `--open-panel`로 패널 자동 열기 |
| `xgen-seepage panel` | 채팅 패널을 기본 브라우저로 연다(`run`이 켜져 있어야 함) |
| `xgen-seepage chat-workflow list`/`set` | 패널 기본 에이전트(워크플로우) 관리(패널 드롭다운으로도 선택 가능) |
| `xgen-seepage status` | 현재 설정·토큰 유효성 확인 |
| `xgen-seepage logout` | 저장된 토큰 삭제 |

## 파이썬 API. 직접 라이브 편집

```python
from xgen_seepage import live_adapter, csv_adapter

# 지금 열려 있는 모든 통합문서 조회
books = live_adapter.list_open_workbooks()
wb_id = books[0].workbook_id

# 시트 구조 파악 (레이아웃, 병합 셀, 수식 존재 여부)
schema = live_adapter.get_sheet_schema(wb_id, sheet=0)
print(schema.preview)

# 셀 하나 실시간으로 갱신. 저장 없이 즉시 화면에 반영됨
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

```python
# Excel이 없는 환경: LibreOffice로 xlsx를 직접 연다(미리 열어둘 필요 없음)
from xgen_seepage import libreoffice_adapter as lo

doc = lo.open_document("report.xlsx")["path"]
schema = lo.get_sheet_schema(doc, sheet=0)
lo.set_cell(doc, sheet=0, row=1, col=2, value="1500")
lo.save(doc)
```

## 노출되는 도구 (24개. `xgen-seepage run`이 XGEN 에이전트 세션에 자동 광고)

**Excel(열려 있는 통합문서)**

| 도구 | 설명 |
|---|---|
| `list_open_workbooks` | 로컬에서 실행 중인 모든 Excel의 열린 통합문서 목록. **항상 첫 호출** |
| `get_live_schema` | 시트 크기·미리보기·병합·수식 존재 여부 |
| `get_live_cell` | 셀 하나의 값+수식+병합 정보 전체(절단 없이) |
| `set_live_cell` | 셀 하나 즉시 갱신(수식 지원, 숫자 자동판정, 병합 anchor 리다이렉트) |
| `read_live_range` / `write_live_range` | 범위 벌크 읽기/쓰기(최대 50,000셀) |
| `append_live_row` | 사용 범위 아래에 행 추가 |
| `activate_live_cell` | Excel 창을 앞으로 + 셀 선택(에이전트가 뭘 보는지 사용자가 눈으로 확인) |

**LibreOffice(Excel 없는 환경, xlsx 파일을 직접 지정해서 연다)**

| 도구 | 설명 |
|---|---|
| `open_libreoffice_document` / `close_libreoffice_document` | 파일 경로로 xlsx를 열고/닫는다 |
| `list_open_libreoffice_documents` | 지금 열려 있는 문서 목록 |
| `get_libreoffice_schema` / `get_libreoffice_cell` | 시트/셀 조회 |
| `set_libreoffice_cell` | 셀 갱신(수식 지원, 병합 anchor 리다이렉트) |
| `read_libreoffice_range` / `write_libreoffice_range` | 범위 벌크 읽기/쓰기 |
| `append_libreoffice_row` | 행 추가 |
| `save_libreoffice_document` | 원본 파일에 저장 |

**CSV**

| 도구 | 설명 |
|---|---|
| `inspect_csv` | CSV 구조 조사(인코딩/구분자 자동 감지 + 미리보기) |
| `get_csv_cell` / `set_csv_cell` | CSV 셀 조회/편집(원본 인코딩·구분자 보존 재저장) |
| `append_csv_row` | CSV 끝에 행 추가 |
| `write_csv_table` | CSV 신규 생성/통째 덮어쓰기 |
| `open_csv_in_excel` | CSV를 로컬 Excel로 열어 라이브 도구로 이어서 편집 |

## stdio MCP 서버로도 쓸 수 있다

기본 경로는 위의 독립 `xgen-seepage run`이지만, 이미 Local MCP 설정이 있는
클라이언트에 그냥 얹고 싶다면 stdio MCP 서버도 남겨뒀다:

```bash
xgen-seepage-mcp
# 또는
python -m xgen_seepage.mcp_server
```

## 개발

```bash
python -m venv .venv
.venv/Scripts/activate       # Windows
pip install -e ".[live,dev]"
pytest tests/ -v
ruff check xgen_seepage/ tests/
```

43개 단위/통합 테스트(인코딩·셀 편집 왕복, MCP dispatcher, 로그인+WS 브릿지
전체 왕복). Excel/LibreOffice가 실제로 열린 상태에서의 셀 IO는 두 애플리케이션이
설치된 머신에서 별도 검증이 필요하다.

## 라이선스

**Apache License 2.0**. [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조. 이
소프트웨어나 그 파생물을 재배포할 때는 저작권·라이선스 고지와 `NOTICE`
파일의 출처 표시를 유지해야 합니다(라이선스 §4).
