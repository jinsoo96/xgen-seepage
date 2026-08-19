# xgen-seepage

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)
[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

**XGEN 에이전트가 당신의 Excel/CSV를 직접 편집하게 해주는 독립형 커넥터.**
ChatGPT나 Claude가 스프레드시트 안에 붙어 셀을 편집하는 것의 XGEN 버전이다.
XGEN 서버에 직접 로그인해 붙는 단일 실행 프로그램이고, 다른 앱을 깔 필요가 없으며,
인터넷 없이 폐쇄망에서도 설치·운영된다.

다루는 형식은 **`.xlsx`** 와 **`.csv`** 다:

- **`.xlsx` (열려 있는 Excel, 실시간)** — 지금 화면에 열어 둔 통합문서를 저장 없이
  그 자리에서 읽고 쓴다(수식 포함). 사용자가 결과를 눈으로 바로 본다.
- **`.xlsx` (Excel 없이)** — Excel이 없는 환경에서도 LibreOffice로 xlsx를 직접 열어
  같은 셀 읽기/쓰기를 제공한다. 파일을 미리 열어둘 필요도 없다.
- **`.csv`** — 인코딩·구분자를 자동 감지해 읽고 편집하며, 필요하면 Excel로 올려 같은
  라이브 경로로 이어 쓴다.

## 지원 환경

| 대상 | 필요 조건 |
|---|---|
| 열려 있는 Excel 실시간 편집(`.xlsx`) | **Windows** 또는 **macOS** + 로컬 Excel |
| Excel 없이 xlsx 편집 | 임의 OS + LibreOffice |
| CSV 편집 | 임의 OS (추가 설치 불필요) |

빌드/설치에만 Python 3.10+ 또는 미리 빌드된 실행파일이 필요하고, 운영은 XGEN 서버
연결만 있으면 된다(폐쇄망 OK).

## 설치

### 방법 A. pip로 설치 (Python 3.10+)

```bash
# 열려 있는 통합문서까지 실시간 편집: [live]
pip install "xgen-seepage[live] @ git+https://github.com/jinsoo96/xgen-seepage.git"

# CSV/LibreOffice만 쓰면 [live] 없이
pip install "xgen-seepage @ git+https://github.com/jinsoo96/xgen-seepage.git"
```

### 방법 B. 단일 실행파일 (Python 불필요)

대상 머신엔 Python이 필요 없다. Python이 있는 빌드 머신에서 한 번 얼려, 그 실행파일
하나만 대상 머신에 복사하면 된다. 폐쇄망이면 인터넷 되는 머신에서 미리 빌드해 반입한다.

```bash
git clone https://github.com/jinsoo96/xgen-seepage.git
cd xgen-seepage
pip install -e ".[live,build]"
python -m PyInstaller packaging/xgen-seepage-connector.spec
```

산출물(레포 루트에서 빌드한 기준):

| OS | 산출물 |
|---|---|
| Windows | `dist/xgen-seepage-connector.exe` |
| macOS | `dist/xgen-seepage-connector` (단일 바이너리) |

> PyInstaller는 크로스컴파일이 안 된다. **대상 OS와 같은 OS에서 빌드**한다(Windows용은
> Windows에서, macOS용은 macOS에서). 라이브 편집 의존성까지 함께 번들된다.

미리 빌드한 실행파일은 [Releases](https://github.com/jinsoo96/xgen-seepage/releases)에
올려 배포할 수 있다.

## 빠른 시작

```bash
xgen-seepage login    # XGEN 서버 URL + 계정 (여러 서버는 목록에서 선택)
xgen-seepage run      # 에이전트-도구 브릿지 + 채팅 패널 상주 (Ctrl+C로 종료)
```

`run`이 떠 있으면, XGEN 어디서(웹 UI 등) 어떤 에이전트와 채팅하든 그 에이전트가 이
머신의 Excel/CSV를 도구로 쓴다. Excel 안에서 바로 채팅하려면:

```bash
xgen-seepage install-excel-addin   # Excel 홈 탭 "XGEN" 그룹에 버튼 설치 (폐쇄망 OK)
# 또는 리본 버튼 없이 브라우저 패널로:
xgen-seepage panel                 # 채팅 패널을 기본 브라우저로 연다
```

패널에서 **에이전트(워크플로우)를 골라** 채팅하면 그 에이전트가 열린 Excel을
편집한다. 사용할 LLM 모델은 그 에이전트플로우의 노드 설정을 그대로 쓴다(패널이
모델을 따로 고르지 않는다). 노드에 모델이 설정돼 있지 않으면 패널이 토스트로 알려준다.

## 동작 방식

```
 [ XGEN 서버 ]  ←(로그인·WebSocket 브릿지)→  [ xgen-seepage run ]  ─→  로컬 Excel / LibreOffice / CSV
      ↑                                             (이 머신)
  에이전트 세션
  (웹 UI·패널 등)
```

1. `run`이 XGEN 서버에 로그인 토큰으로 붙어, 이 프로세스가 가진 도구 카탈로그를
   에이전트-도구 WebSocket 브릿지로 광고한다.
2. 사용자가 XGEN에서 에이전트와 채팅하다 도구를 호출하면, 그 호출이 브릿지를 타고
   이 머신으로 와서 Excel/CSV를 실제로 편집한다.
3. 셀 편집은 로컬 프로세스 간 통신으로만 일어난다(웹서버·Office 애드인 인프라
   불필요). 채팅 패널도 XGEN 로그인에만 의존하는 순수 웹 UI다.

## 운영

**여러 XGEN(jeju/dev/prod) 전환**: `login` 때 써 본 서버 목록에서 고르고, 토큰은
서버별로 따로 저장된다. `xgen-seepage server list`로 확인하고 `server use`로 전환한다
(그 서버에 토큰이 있으면 재로그인 없이). 지금 붙은 서버·계정은 패널 헤더와
`xgen-seepage status`에서 확인된다 — 권한 부족으로 나는 403을 "엉뚱한 서버·계정"으로
바로 진단할 수 있다.

**폐쇄망 TLS(내부 CA)**: 사내 내부 CA로 발급한 HTTPS 인증서를 쓰는 XGEN이라도, 그 CA가
PC의 OS 신뢰 저장소에 설치돼 있으면(폐쇄망은 GPO로 보통 이미 설치돼 있다) 별도 설정
없이 그대로 붙는다. certifi 고정 목록이 아니라 OS 신뢰 저장소로 검증하기 때문이다.
CA가 OS에 없는 예외적 경우에만 `login --allow-private-certificate`로 검증을 끌 수 있다
(그 서버 한정, 권장하지 않음).

**비대화형/대량 설치**: 환경변수 `XGEN_SEEPAGE_SERVER_URL` / `XGEN_SEEPAGE_EMAIL` /
`XGEN_SEEPAGE_PASSWORD`로 `login`을 스크립트로 자동화할 수 있다.

**플랫폼별 참고**:
- **macOS** — 열린 Excel 편집은 자동화 권한이 한 번 필요하다. 첫 실행 시 "Excel을
  제어하려 합니다" 창에서 허용하거나 [시스템 설정 > 개인정보 보호 및 보안 > 자동화]에서
  켠다. 단일 바이너리는 ad-hoc 서명이라 Gatekeeper가 막으면 최초 1회 우클릭 > 열기.
- **Windows** — conda로 빌드하면 스펙이 conda의 OpenSSL DLL까지 번들한다(안 그러면 얼린
  exe가 `_ssl` 로드 실패로 XGEN에 못 붙는다). 일반 venv에선 무관.

## CLI

| 명령 | 설명 |
|---|---|
| `xgen-seepage login` | XGEN 서버에 로그인(써 본 서버 목록에서 선택), 토큰을 서버별로 저장 |
| `xgen-seepage run` | 에이전트-도구 브릿지 + 채팅 패널 로컬 서버 상주(포그라운드). `--open-panel`로 패널 자동 열기 |
| `xgen-seepage server list`/`use` | 여러 XGEN 확인·전환(그 서버에 토큰이 있으면 재로그인 없이) |
| `xgen-seepage panel` | 채팅 패널을 기본 브라우저로 연다(`run`이 켜져 있어야 함) |
| `xgen-seepage install-excel-addin` | Excel 리본에 채팅 버튼 설치(폐쇄망 OK). `uninstall-excel-addin`으로 제거 |
| `xgen-seepage chat-workflow list`/`set` | 패널 기본 에이전트(워크플로우) 관리(패널 드롭다운으로도 선택 가능) |
| `xgen-seepage status` | 현재 설정·토큰 유효성·권한 확인 |
| `xgen-seepage logout` | 저장된 토큰 삭제 |

## 파이썬 API (직접 라이브 편집)

```python
from xgen_seepage import live_adapter, csv_adapter

# 지금 열려 있는 모든 통합문서 조회
books = live_adapter.list_open_workbooks()
wb_id = books[0].workbook_id

# 시트 구조 파악(레이아웃·병합 셀·수식 존재 여부)
schema = live_adapter.get_sheet_schema(wb_id, sheet=0)
print(schema.preview)

# 셀 하나 실시간 갱신(저장 없이 즉시 화면 반영). 수식도 그대로.
live_adapter.set_cell(wb_id, sheet=0, row=1, col=2, value="1500")
live_adapter.set_cell(wb_id, sheet=0, row=1, col=3, value="=C2*1.1", as_formula=True)

# 범위 벌크 읽기(값 + 수식 둘 다)
data = live_adapter.read_range(wb_id, sheet=0, row0=0, col0=0, row1=10, col1=5)
```

```python
# CSV: 인코딩/구분자 자동 감지. 필요하면 Excel로 승격
table = csv_adapter.load_table("sales_2026.csv")
csv_adapter.set_cell("sales_2026.csv", row=1, col=1, value="9800")
wb = csv_adapter.open_in_excel("sales_2026.csv")

# Excel이 없는 환경: LibreOffice로 xlsx 직접 열기
from xgen_seepage import libreoffice_adapter as lo
doc = lo.open_document("report.xlsx")["path"]
lo.set_cell(doc, sheet=0, row=1, col=2, value="1500")
lo.save(doc)
```

## 노출되는 도구 (57개, `xgen-seepage run`이 XGEN 에이전트 세션에 자동 광고)

**Excel(열려 있는 통합문서)** — 조회

| 도구 | 설명 |
|---|---|
| `list_open_workbooks` | 실행 중인 모든 Excel의 열린 통합문서 목록. **항상 첫 호출** |
| `get_workbook_overview` | 모든 시트의 크기·미리보기·수식여부 한 번에. 복잡한 파일 파악의 시작점 |
| `get_live_schema` | 시트 크기·미리보기·병합·수식 존재 여부 |
| `get_live_cell` | 셀 하나의 값+수식+병합 정보 전체(절단 없이) |
| `read_live_range` | 범위 벌크 읽기(값+수식, 최대 50,000셀) |
| `get_live_table_region` | 특정 셀이 속한 표의 실제 범위(테두리·빈 행/열 기준) |

**Excel** — 값/수식 편집

| 도구 | 설명 |
|---|---|
| `set_live_cell` | 셀 하나 즉시 갱신(수식 지원, 숫자 자동판정, 병합 anchor 리다이렉트) |
| `write_live_range` | 범위 벌크 쓰기 |
| `append_live_row` | 사용 범위 아래에 행 추가 |
| `find_replace_live` | 사용 범위 전체 찾기/바꾸기 |
| `recalculate_live` | 모든 수식 강제 재계산 |
| `activate_live_cell` | Excel 창을 앞으로 + 셀 선택(진행을 눈으로 확인) |

**Excel** — 구조 편집(원본 비파괴, 수식·서식째 이동)

| 도구 | 설명 |
|---|---|
| `insert_live_rows` / `delete_live_rows` | 행 삽입 / 삭제 |
| `insert_live_columns` / `delete_live_columns` | 열 삽입 / 삭제 |
| `copy_live_range` | 범위 복사(같은/다른 시트, 수식 상대참조 자동 조정) |
| `set_live_rows_visible` / `set_live_columns_visible` | 행 / 열 숨기기·보이기 |
| `sort_live_range` | 특정 열 기준 정렬(Excel 네이티브, 행 전체 서식·수식 보존) |
| `set_live_autofilter` | 자동필터 켜기/끄기(**Windows 전용**) |

**Excel** — 시트

| 도구 | 설명 |
|---|---|
| `list_live_sheets` | 시트 목록(순서·이름) |
| `add_live_sheet` / `rename_live_sheet` / `delete_live_sheet` | 시트 추가 / 이름변경 / 삭제 |
| `move_live_sheet` | 시트 위치 이동(재정렬, 서식·수식 보존) |
| `define_live_name` / `list_live_names` / `delete_live_name` | 이름 정의 추가 / 목록 / 삭제 |

**Excel** — 서식/레이아웃

| 도구 | 설명 |
|---|---|
| `set_live_fill_color` | 셀 범위 배경색('노란색'/'#FFFF00') |
| `color_live_rows_where` | **조건부 행 강조** — '어떤 열=값'인 행을 한 번에 색칠(행 안 짚어 어긋남 방지) |
| `format_live_range` | 채우기색·굵게·글자색 한 번에 |
| `set_live_number_format` | 표시 형식(천단위·백분율·날짜·통화). **값은 안 바꿈** |
| `set_live_borders` | 테두리(격자/바깥/개별 변, thin·medium·thick) |
| `set_live_wrap_text` | 셀 자동 줄바꿈(긴 내용 다 보이게, 행 높이 자동 맞춤) |
| `merge_live_cells` / `unmerge_live_cells` | 셀 병합 / 해제 |
| `freeze_live_panes` | 틀 고정(상단 N행 / 좌측 N열) |
| `autofit_live` | 열 너비·행 높이를 내용에 맞춤 |
| `set_live_column_width` / `set_live_row_height` | 열 너비 / 행 높이 지정 |
| `set_live_data_validation` | 드롭다운(목록) 유효성 검사 |
| `clear_live_range` | 범위의 값/서식 지우기 |

**LibreOffice(Excel 없는 환경, xlsx 파일 경로를 직접 지정해 연다)**

| 도구 | 설명 |
|---|---|
| `open_libreoffice_document` / `close_libreoffice_document` | 파일 경로로 xlsx 열고/닫기 |
| `list_open_libreoffice_documents` | 열려 있는 문서 목록 |
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

기본 경로는 위의 독립 `xgen-seepage run`이지만, 이미 로컬(stdio) MCP를 지원하는
클라이언트에 그냥 얹고 싶다면 stdio MCP 서버도 제공한다. [`examples/local_mcp_stdio.md`](examples/local_mcp_stdio.md) 참조.

```bash
xgen-seepage-mcp        # 또는 python -m xgen_seepage.mcp_server
```

## 개발

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e ".[live,dev]"
ruff check xgen_seepage/
mypy xgen_seepage/
```

## 라이선스

**Apache License 2.0**. [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조. 이 소프트웨어나
그 파생물을 재배포할 때는 저작권·라이선스 고지를 유지해야 합니다(라이선스 §4).
