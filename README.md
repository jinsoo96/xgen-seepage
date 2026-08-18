# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)

**폐쇄망에서도 설치되는 독립형 XGEN Excel/CSV 커넥터.** 챗지피티나 클로드가
엑셀 안에 플러그인처럼 붙어서 편집·붙여넣기하는 것의 XGEN 버전. XGEN 서버에
직접 로그인해서 붙는 **독립 실행 프로그램**이고(다른 앱 설치 불필요), 인터넷
없이도 설치·운영 가능하다.

- **로그인 한 번**(`xgen-seepage login`) → XGEN 서버·계정에 붙는다. 여러
  XGEN(jeju/dev/prod)을 오갈 땐 로그인 시 써 본 서버 목록에서 고르고, 토큰은
  서버별로 따로 저장한다(`xgen-seepage server use`로 재로그인 없이 전환).
  패널·`status`가 지금 붙은 서버와 계정을 보여줘, 권한 부족으로 나는 403을
  "엉뚱한 서버·계정" 문제로 바로 진단할 수 있다.
- **`xgen-seepage run`을 상주**시키면, 사용자가 XGEN 어디서 어떤 에이전트와
  채팅하든 그 에이전트가 사용자 로컬 Excel/CSV를 도구로 쓸 수 있게 된다.
- 지금 로컬 Excel에서 **열려 있는** 통합문서를 실시간으로 읽고 쓴다(수식 포함,
  파일 저장 불필요. 사용자가 화면에서 바로 결과를 본다). 별도 프로세스로 도는
  커넥터가 사용자가 연 통합문서를 찾아 에이전트가 셀을 실제로 바꾼다.
- **Excel 안 채팅 패널**: `xgen-seepage panel`이 브라우저로 채팅 패널을 연다.
  패널에서 **에이전트(워크플로우)와 모델(provider)을 골라** 채팅하면 그
  에이전트가 열린 Excel을 편집한다. 이 패널은 office.js에 의존하지 않는 순수
  웹 UI라 XGEN 로그인에만 의존한다(인터넷·Microsoft 애드인 인프라 불필요).
  Excel 리본 버튼으로도 띄울 수 있다 - `xgen-seepage install-excel-addin`
  한 번이면 홈 탭 "XGEN" 그룹에 "xgen-seepage" 버튼이 뜬다(**폐쇄망 OK**,
  Microsoft 클라우드 불필요한 로컬 레지스트리 사이드로드).
- **Excel이 없는 환경**(폐쇄망 등)에서도 LibreOffice로 xlsx를 직접 열어 같은
  셀 읽기/쓰기를 제공한다. 파일을 미리 열어둘 필요도 없다.
- CSV는 인코딩/구분자를 자동 감지해 읽고 편집하며, 필요하면 Excel로 올려 같은
  라이브 경로로 넘어간다.

## 설치 & 사용

**필요한 것**: Python 3.10+. 열려 있는 Excel을 실시간 편집하려면 Windows/macOS +
로컬 Excel. Excel이 없어도 CSV와 (LibreOffice가 있으면) xlsx 편집은 된다.
인터넷은 설치 때만 필요하고, 운영은 XGEN 서버 연결만 있으면 된다(폐쇄망 OK).

**macOS**: 열린 Excel 편집(라이브)은 macOS 자동화 권한이 한 번 필요하다. 처음
`xgen-seepage run` 후 에이전트가 Excel을 만지려 하면 "터미널이 Microsoft Excel을
제어하려 합니다" 창이 뜨는데 '확인'을 누르면 된다(또는 [시스템 설정 > 개인정보
보호 및 보안 > 자동화]에서 직접 허용). Windows의 COM처럼 Mac은 AppleScript로
붙는데, 이 권한 없이는 OS가 막는다.

**어느 머신에서든 설치 (GitHub에서 바로):**

```bash
# 라이브(열려 있는 통합문서) 편집까지: [live], CSV만 쓰면 뒤의 [live] 없이
pip install "xgen-seepage[live] @ git+https://github.com/jinsoo96/xgen-seepage.git"

xgen-seepage login    # XGEN 서버 URL + 계정 (여러 서버는 목록에서 선택)
xgen-seepage run      # 에이전트-도구 브릿지 + 채팅 패널 상주 (Ctrl+C로 종료)
```

`run`이 떠 있으면, XGEN 어디서(웹 UI 등) 어떤 에이전트와 채팅하든 그 에이전트가
이 머신의 Excel/CSV를 도구로 쓴다. Excel 안에서 바로 쓰려면:

```bash
xgen-seepage install-excel-addin   # Excel 홈 탭 "XGEN" 그룹에 버튼 설치(폐쇄망 OK)
# 또는 리본 버튼 없이 브라우저로:
xgen-seepage panel                 # 채팅 패널을 기본 브라우저로 연다
```

**여러 XGEN(jeju/dev/prod)을 오갈 때**: `login` 때 써 본 서버 목록에서 고르고,
토큰은 서버별로 저장된다. `xgen-seepage server list`로 확인하고 `server use`로
전환한다(그 서버에 토큰이 있으면 재로그인 없이). 지금 붙은 서버·계정은 패널
헤더와 `xgen-seepage status`에서 확인된다.

**파이썬 없이 단일 실행파일로 (Windows·macOS)**: 대상 머신엔 파이썬이 필요 없다.
빌드 머신(파이썬 있는 곳)에서 한 번 얼려서 그 실행파일 하나만 대상에 반입/복사하면
`login`/`run`/`status`/`server list`/`install-excel-addin` 전부 그대로 돈다.
폐쇄망이면 인터넷 있는 머신에서 미리 얼려 반입하면 된다.

```bash
git clone https://github.com/jinsoo96/xgen-seepage.git
cd xgen-seepage
pip install -e ".[live,build]"
python -m PyInstaller packaging/xgen-seepage-connector.spec
# 산출물(스펙을 레포 루트에서 돌린 기준):
#   Windows: dist/xgen-seepage-connector.exe
#   macOS:   dist/xgen-seepage-connector   (Mach-O 단일 바이너리)
```

- **플랫폼별로 빌드해야 한다**(PyInstaller는 크로스컴파일이 안 됨): Windows용
  exe는 Windows에서, macOS용 바이너리는 macOS에서 얼린다. 대상 OS와 같은 OS에서
  빌드. 라이브(열린 Excel) 편집 의존성(Windows=pywin32, macOS=appscript)까지 함께
  번들된다 - 실측으로 두 플랫폼 다 얼린 실행파일만으로 XGEN 로그인(HTTPS)이 되는 걸
  확인했다.
- **conda 빌드 함정(Windows)**: 빌드 머신이 conda면 스펙이 conda의 OpenSSL DLL
  (`Library\bin`)까지 챙겨 번들한다. 안 그러면 얼린 exe가 `_ssl` DLL 로드 실패로
  XGEN에 못 붙는다(일반 venv·macOS에선 자동으로 무관).
- **macOS 첫 실행**: ad-hoc 서명이라 Gatekeeper가 막으면 최초 1회 우클릭 > 열기.
  키체인·Excel 자동화 권한은 실행 시 macOS가 물어보면 허용.

**폐쇄망 TLS(내부 CA)**: 폐쇄망 XGEN이 사내 내부 CA로 발급한 HTTPS 인증서를
쓰더라도, 그 CA가 PC의 OS 신뢰 저장소에 설치돼 있으면(폐쇄망에선 GPO로 보통
이미 설치돼 있다) **별도 설정 없이 그대로 붙는다**. 커넥터가 certifi 고정
목록이 아니라 OS 신뢰 저장소로 인증서를 검증하기 때문이다(`truststore`).
CA가 OS에 안 깔려 있는 예외적 경우에만 `login --allow-private-certificate`로
검증을 끌 수 있다(그 서버 한정, 권장하지 않음).

**비대화형/대량 설치**(로그인을 스크립트로): 환경변수 `XGEN_SEEPAGE_SERVER_URL` /
`XGEN_SEEPAGE_EMAIL` / `XGEN_SEEPAGE_PASSWORD`로 `login`을 자동화할 수 있다.

## CLI

| 명령 | 설명 |
|---|---|
| `xgen-seepage login` | XGEN 서버에 로그인(써 본 서버 목록에서 선택), 토큰을 서버별로 OS 키체인에 저장 |
| `xgen-seepage run` | 에이전트-도구 브릿지 + 채팅 패널 로컬 서버를 상주시킨다(포그라운드). `--open-panel`로 패널 자동 열기 |
| `xgen-seepage server list`/`use` | 여러 XGEN(jeju/dev/prod) 확인·전환. 그 서버에 토큰이 이미 있으면 재로그인 없이 전환 |
| `xgen-seepage panel` | 채팅 패널을 기본 브라우저로 연다(`run`이 켜져 있어야 함) |
| `xgen-seepage install-excel-addin` | Excel 리본에 XGEN 채팅 버튼 설치(**폐쇄망 OK**, 클라우드 불필요). `uninstall-excel-addin`으로 제거 |
| `xgen-seepage chat-workflow list`/`set` | 패널 기본 에이전트(워크플로우) 관리(패널 드롭다운으로도 선택 가능) |
| `xgen-seepage status` | 현재 설정·토큰 유효성·권한 확인(권한이 없으면 패널에서 403이 날 수 있음을 미리 알려줌) |
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

## 노출되는 도구 (37개. `xgen-seepage run`이 XGEN 에이전트 세션에 자동 광고)

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
| `list_live_sheets` | 시트 목록(순서·이름) |
| `add_live_sheet` | 새 시트 추가(이름·위치 지정 가능, 기본은 맨 뒤) |
| `rename_live_sheet` / `delete_live_sheet` | 시트 이름변경 / 삭제 |
| `move_live_sheet` | 시트 위치 이동(재정렬, Excel 네이티브 - 서식·수식 보존) |
| `set_live_fill_color` | 셀 범위 배경색 칠하기(조건부 행 강조 등). '노란색'/'#FFFF00' |
| `format_live_range` | 채우기색·굵게·글자색 한 번에 |
| `set_live_number_format` | 표시 형식(천단위·백분율·날짜·통화 등). **값은 안 바꿈** |
| `merge_live_cells` / `unmerge_live_cells` | 셀 병합 / 해제 |
| `autofit_live` | 열 너비·행 높이를 내용에 맞춤 |
| `set_live_column_width` / `set_live_row_height` | 열 너비 / 행 높이 지정 |

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
ruff check xgen_seepage/
mypy xgen_seepage/
```

## 라이선스

**Apache License 2.0**. [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조. 이
소프트웨어나 그 파생물을 재배포할 때는 저작권·라이선스 고지와 `NOTICE`
파일의 출처 표시를 유지해야 합니다(라이선스 §4).
