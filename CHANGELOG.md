# Changelog

## 0.21.0 (2026-08-19)

**라이브 도구 API 재설계 - 파라미터가 균질한 조작을 응집도 높은 액션 도구로.**
도구 수를 무작정 줄이는 임시방편(모드 토글) 대신, 성격이 같은 조작을 하나의
action 도구로 묶어 스키마를 깔끔하게 재구성했다. 노출 도구 57 → 40개. 모든
액션을 macOS에서 call_tool 경유로 실동작 검증.

- `format_live_range(action=…)`: 범위 서식 통합 - fill_color / font / number_format /
  borders / wrap / merge / unmerge / data_validation / clear (개별 도구 9개를 흡수).
- `edit_live_rows_columns(action=…)`: 행·열 삽입/삭제 (4개 흡수).
- `set_live_visibility(action=…)`: 행·열 숨김/표시 (2개 흡수).
- `resize_live(action=…)`: 열너비/행높이 (2개 흡수).
- `manage_live_sheet(action=…)`: 시트 목록/추가/이름변경/삭제/이동 (5개 흡수).
- `manage_live_name(action=…)`: 이름정의 목록/추가/삭제 (3개 흡수).
- 성격이 뚜렷한 도구(조회·값편집·color_rows_where·copy·freeze·autofilter·autofit·
  activate·find_replace·recalculate)는 단일 도구로 그대로 유지.
- **누락돼 있던 `sort_live_range`(정렬)·범위 clear를 정식 노출** — live_adapter엔
  구현돼 있었으나 도구로 광고되지 않던 것. sort는 단일 도구, clear는
  format_live_range의 action으로 노출.

## 0.20.0 (2026-08-19)

**긴 셀 내용 다 보이게 + 표 범위 인식.** 실무 표를 다룰 때 자주 필요한 두 가지.
전부 macOS/Windows 양쪽 실측. 도구 55 → 57개.

- `set_live_wrap_text` 신규: 셀 자동 줄바꿈. 내용이 길어 잘려 보일 때 켜면
  셀 안에서 줄바꿈돼 다 보이고, autofit_rows로 행 높이도 내용에 맞춰 늘린다
  (열 너비는 안 건드림. 열을 넓혀 한 줄로 보이게 하려면 기존 autofit_live).
- `get_live_table_region` 신규: 지정한 셀이 속한 표의 실제 경계를 반환한다.
  Excel의 현재 영역(빈 행·열로 둘러싸인 연속 블록)을 기준으로 해, 값·테두리·
  서식으로 그린 표의 범위를 정확히 잡는다. 참고로 Excel의 used_range 자체가
  테두리만 있는 빈 셀까지 포함하므로, 스키마의 rows/cols도 이미 테두리로
  그린 표 범위를 반영한다(실측 확인).

## 0.19.0 (2026-08-18)

**"엑셀 내부 기능이 지시하에 다 돌게."** 셀 편집을 넘어 실무에서 쓰는 통합문서
조작 전반을 도구로 노출. 전부 macOS(appscript)/Windows(COM) 양쪽에서 **실측**한
방법만 넣었다(추측 배제). 총 도구 38 → 55개.

- **조건부 행 강조** `color_live_rows_where` 신규: "어떤 열의 값이 X인 행"을
  에이전트가 행 번호를 손으로 짚지 않고 조건으로 한 번에 색칠. 헤더가 여러 줄인
  실무 대장에서 행이 어긋나거나 일부만 칠해지던 사고(에이전트가 `set_live_fill_color`
  를 행마다 호출하다 오프셋으로 틀림)를 구조적으로 차단. clear_non_matching으로
  잘못 칠해진 것도 함께 정리. `set_live_fill_color` 설명에 "조건부 색칠은 이걸 써라" 명시.
- **구조 편집**: `insert_live_rows`/`delete_live_rows`/`insert_live_columns`/
  `delete_live_columns`(수식·서식째 밀림, 원본 비파괴), `copy_live_range`(같은/다른
  시트, 수식 상대참조 자동 조정), `set_live_rows_visible`/`set_live_columns_visible`.
- **서식/레이아웃**: `set_live_borders`(격자/바깥/개별 변, thin·medium·thick),
  `freeze_live_panes`(틀 고정/해제).
- **데이터**: `find_replace_live`(사용 범위 찾기/바꾸기), `set_live_data_validation`
  (드롭다운 목록), `set_live_autofilter`(자동필터 - Windows 전용, macOS는 미지원이라
  정렬·조건부강조로 대체하도록 친절 에러), `recalculate_live`(수식 강제 재계산).
- **이름 정의**: `define_live_name`/`list_live_names`/`delete_live_name`.
- **macOS 한글 NFD 매칭 버그 수정**: macOS 파일시스템이 한글 파일·시트명을 NFD로
  보관해 NFC 리터럴과 매칭이 조용히 빗나가던 문제(`_sheet`/`_resolve_book`에서
  `unicodedata.normalize("NFC")` 재시도).
- **토큰 저장 파일 폴백**: OS 키체인에 접근할 수 없는 헤드리스/서비스/원격(SSH)/
  폐쇄망 환경에서 로그인이 막히던 문제를, 키체인 실패 시 설정 폴더
  `tokens.json`(권한 0600)으로 폴백해 해결. 키체인 되는 데스크톱에선 키체인 우선.
- 브릿지가 도구 카탈로그 동기화 시 "보낸 N개, 서버 등록 N개"를 로그로 남기게(관측성).

## 0.18.0 (2026-08-18)

**통합문서 전체를 읽고, 그 안에서 원본 보존하며 자유롭게.** "엑셀에 침투했으면
파일·데이터 구조를 다 읽고 그 안에서 에이전트가 자유롭게 활동해야 한다"는 요구.

- `get_workbook_overview` 신규: **모든 시트의 크기·미리보기·수식여부·병합을 한
  번에** 반환. 시트 여러 개짜리 복잡한 파일을 에이전트가 한 눈에 파악하는
  시작점. 실측: 3시트(매출·집계·메모) 전체 구조 + 크로스시트 수식(=SUM(매출!..))
  계산값까지 한 번에 확인.
- **비파괴 원칙을 도구 설명에 못박음**: `list_open_workbooks`(첫 호출 도구)에
  "원본 보존, 요청한 부분만 정확히, 통째로 지우거나 다시 쓰지 마라"를 명시.
  `write_live_range`엔 "지정 범위의 기존 값·수식·서식이 사라진다, 바꿀 셀만
  좁게, 몇 개면 set_live_cell을 써라" 경고. `delete_live_sheet`는 "명시적 요청
  시에만". 에이전트가 surgical하게 편집하도록 유도.

## 0.17.0 (2026-08-18)

**서식/구조 도구 확장 + 비파괴 원칙.** "엑셀에 침투했으면 전체 기능을 쓸 수
있어야 한다"는 요구에 맞춰 서식·레이아웃 도구를 넓혔고, 무엇보다 **원본을
갈아엎지 않고 지정한 것만 고치는(surgical) 비파괴**를 원칙으로 못박았다.

- 색/서식: `set_live_fill_color`(배경색, 조건부 행 강조), `format_live_range`
  (채우기·굵게·글자색), `set_live_number_format`(표시 형식만, 값 불변),
  `merge_live_cells`/`unmerge_live_cells`, `autofit_live`,
  `set_live_column_width`/`set_live_row_height`. (도구 29 → 37)
- **비파괴 검증(실측, 별도 워크북)**: 수식이 든 표에 숫자서식·병합·자동맞춤·
  크기 조정을 적용해도 수식·다른 셀이 그대로 보존됨을 확인.
- **정렬을 네이티브로**: 초기 `sort_range`가 값을 다시 써넣어 수식/서식을
  날리던 걸(비파괴 위반) Excel 네이티브 정렬로 바꿔 행 전체를 서식·수식까지
  보존하며 재배치하게 했다.
- 읽기 도구(`read_live_range`/`get_live_cell`/`get_live_schema`)는 값과 함께
  **수식**을 준다 - 에이전트가 "여긴 수식이 있으니 함부로 덮지 말자"를 인지한다.

## 0.16.0 (2026-08-18)

**시트 관리 도구 추가.** 실사용에서 나온 구멍: 에이전트한테 "시트 하나 만들어"/
"시트 옮겨" 하면 못 했다 - 셀 편집 도구만 있었지 시트 자체를 다루는 도구가
없어서 에이전트가 "문서 생성해드릴까요" 같은 헛소리로 빠졌다. 시트 도구 5개를
추가했다(총 24 → 29개).

- `list_live_sheets`(목록), `add_live_sheet`(추가, 이름·위치 지정 가능·기본 맨 뒤),
  `rename_live_sheet`, `delete_live_sheet`, `move_live_sheet`(재정렬).
- 크로스플랫폼(Windows COM=`Move(Before/After)`, macOS appscript=`move(to=...)`).
  실측(실제 Mac Excel): 추가·이름변경·이동(맨 앞으로)·삭제 전부 정상.

## 0.15.0 (2026-08-18)

**macOS 지원 - 실제 Mac(Apple Silicon, Excel 설치)에서 세팅하며 이슈를 잡았다.**
지금까지 Windows에서만 검증됐고 macOS 경로는 "미검증"이었는데, 실제 Mac에
설치해 보니 Windows 전용 코드가 여러 곳 있었다. OS별로 갈라 고치고, 고칠 수
있는 건 그 Mac에서 실측했다.

- **인증서 신뢰**: `certs.py`가 `certutil`(Windows)뿐이었다 → macOS는
  `security add-trusted-cert`로 로그인 키체인에 신뢰 등록. 실측: 인증서가
  로그인 키체인에 실제 등록됨.
- **리본 애드인 설치**: `install-excel-addin`이 Windows 레지스트리 전용이라
  Mac에선 "윈도우 전용"으로 빠졌다 → macOS는 매니페스트를
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`에 사이드로드.
  실측: 매니페스트가 올바른 라벨로 그 폴더에 설치됨.
- **LibreOffice 탐색**: `soffice.exe`(Windows 경로)만 찾았다 → macOS
  (`/Applications/LibreOffice.app/Contents/MacOS/soffice`)·Linux 경로 추가.
- **live_adapter(열린 Excel 편집)**: COM 전용 코드를 크로스플랫폼으로. `saved`
  판정을 방어적으로(COM `Saved` → appscript `saved` → 안전 폴백), 병합 감지는
  COM 속성 의존이라 non-Windows에선 건너뛴다(일반 셀 읽기/쓰기는 그대로).
- **macOS 자동화 권한(TCC) 안내**: Mac은 Excel 제어에 자동화 권한이 필요한데,
  없으면 xlwings가 raw appscript 오류(-1743)로 죽었다 → "[시스템 설정 > 개인정보
  보호 및 보안 > 자동화]에서 Excel 제어를 허용하세요"라는 실행 가능한 메시지로
  바꿈. 실측: 권한 없는 상태에서 그 안내가 뜬다.
- 로그인/TLS(truststore OS 신뢰)도 실제 Mac에서 dev-xgen 200 OK로 확인.

## 0.14.3 (2026-08-17)

**폐쇄망 내부 CA 대응 - OS 신뢰 저장소로 TLS 검증(설치파일 활용의 핵심).**
"설치파일로 들고 가 폐쇄망에서 쓴다"는 관점으로 다시 파보니 TLS가 진짜
문제였다. 시스템 인증서를 인지하는 클라이언트는 폐쇄망 XGEN에 그대로 붙는데,
우리 커넥터는 httpx 기본값인 `certifi`(고정 공개 CA 목록)만 봐서 **OS 신뢰
저장소를 안 봤다**.
그래서 폐쇄망 내부 CA가 PC(OS)에 설치돼 있어도 XGEN 인증서를 거부해 접속이
실패하고, 유일한 우회가 검증 전면 해제뿐이었다.

- HTTP(httpx)·WebSocket 브릿지·태스크팬 프록시 모두 **OS 신뢰 저장소로 검증**
  하도록 바꿈(`truststore`, 없으면 stdlib 기본 컨텍스트로 폴백. WS는
  websockets가 이미 OS 저장소를 봤다). 내부 CA가 OS에 깔려 있으면(폐쇄망 GPO
  기본) 별도 설정 없이 그대로 붙는다. `--allow-private-certificate`(검증 해제)는
  CA가 OS에 없는 예외적 최후 수단으로만.
- 실측: OS 신뢰 컨텍스트(truststore·stdlib 폴백)로 실제 dev-xgen(공개 인증서)에
  TLS 핸드셰이크 성공(회귀 없음). `truststore` 의존성 추가 + 스펙 hiddenimport.

## 0.14.2 (2026-08-17)

**얼린 exe가 실행이 안 되던 진짜 버그 수정(재검수로 발견).** 공개된 코드로
exe를 다시 얼려보니 첫 HTTPS 관련 import에서 `ImportError: DLL load failed
while importing _ssl`로 즉사했다 - 즉 폐쇄망 설치파일이 XGEN에 아예 못 붙는
상태였다. 원인: conda Python은 OpenSSL DLL을 `<env>\Library\bin`에 두는데
PyInstaller가 거길 안 뒤져 `libssl-3-x64.dll`/`libcrypto-3-x64.dll`이 번들에서
빠졌고, `_ssl.pyd`가 시스템의 잘못된 버전을 물어 "procedure not found"가 났다.

- 스펙이 conda OpenSSL DLL을 명시적으로 번들하게 수정(일반 venv에선 no-op).
  재빌드한 exe로 `--help`/`server list`/`status`와 실제 `login`(HTTPS 200 OK)까지
  얼린 exe로 재확인.
- README의 exe 산출물 경로를 실제(`dist/`)로 정정하고 conda 빌드 주의를 추가.

## 0.14.1 (2026-08-17)

**설치 안내 + 리본 버튼 브랜딩.**

- README에 "어느 머신에서든 설치" 가이드 추가(GitHub에서 바로 pip install,
  login/run/install-excel-addin/panel, 서버 전환, 폐쇄망 exe 빌드). 다른
  서버에도 바로 따라 설치할 수 있게 자립형으로.
- 리본 버튼 아이콘을 XGEN 로고로, 라벨을 "xgen-seepage 채팅" -> "xgen-seepage"로.
  매니페스트 Version/Id를 올려 Office가 캐시를 버리고 새로 읽게 함(라벨/아이콘을
  바꾸면 Version을 올려야 반영됨).
- 내부 설계노트 `ARCHITECTURE.md`는 공개 레포에서 빼고 로컬 보존(gitignore).

## 0.14.0 (2026-08-17)

**서버 전환(jeju/dev/prod) + 403 Forbidden을 실행 가능한 안내로.** 여러
XGEN을 오갈 때 토큰이 어긋나 403이 나는 걸 구조적으로 없애고, 권한 부족으로
나는 403을 "뭘 해야 하는지" 알려주게 했다. 리본 버튼 → 도킹 패널 → 채팅 →
실제 셀 편집(A1→42) 전 경로를 실서버로 재검증했다(`ARCHITECTURE.md` §11).

- **토큰을 서버별로 저장**: 예전엔 단일 슬롯이라 서버 B로 바꿔도 키체인엔
  A의 토큰이 남아 B에 A 토큰을 보내 403이 났다. 키를 `access_token::<server>`
  로 나눠 그 불일치 자체를 없앴다(예전 단일 슬롯은 자동 마이그레이션). 로그인
  순서도 고쳐, 토큰 저장이 실패하면 활성 서버를 바꾸지 않는다.
- **`xgen-seepage server list`/`use` 신규**: 로그인해 본 서버와 토큰 보유
  여부를 보고, 그 서버에 토큰이 있으면 재로그인 없이 활성 서버를 전환한다
  (여러 서버를 오갈 때 "서버 전환=계정 공간 전환"이 되도록 서버별 토큰으로 관리).
- **403을 명확히**: 403은 토큰이 아니라 권한 문제다(게이트웨이가 JWT로 판정).
  패널과 `status`가 "이 계정은 이 서버에서 권한이 없다 - 서버/계정을 확인하라"
  고 알려준다. 패널 헤더에 지금 붙은 서버·계정을 표시(`/server`)해 엉뚱한
  서버에 붙은 걸 눈으로 잡게 했다.
- 실행 요청 바디를 XGEN 실행 API 형식(`include_tool_events` 등)에 맞춤.
- 저장소 정리: 테스트·내부 worklog를 공개 레포에서 빼고 로직만 남김.

## 0.13.0 (2026-08-17)

**Excel 리본 버튼을 폐쇄망에서 실제로 띄우는 데 성공 + `install-excel-addin`.**
전에 "이 Office 볼륨 빌드는 웹 애드인을 못 띄운다"고 한 결론이 틀렸다.
사용자 지적("M365는 폐쇄망이면 어차피 안 되잖아")이 맞았고, 클라우드를
안 거치는 로컬 개발자 레지스트리 사이드로드를 제대로 파고들었다.
자세한 내용은 `ARCHITECTURE.md` §11.

- **원인**: `HKCU\...\WEF\Developer` 레지스트리 값의 **이름**도 매니페스트
  전체 경로여야 하는데(이름=경로, 데이터=경로) 이름을 다른 문자열로 넣어서
  Office가 안 읽었던 것. 이름을 경로로 고치니 실제 Excel 홈 탭에
  "XGEN > xgen-seepage 채팅" 버튼이 떴다(스크린샷 확인). 완전 로컬,
  클라우드·인터넷 불필요 - 폐쇄망 호환.
- `xgen-seepage install-excel-addin` 신규: 인증서 생성+신뢰 + 매니페스트
  로컬 설치(포트 패치) + 레지스트리 등록을 한 명령으로. 실측: 명령이
  올바른 사이드로드를 등록하는 것 확인. `uninstall-excel-addin`도 추가.
- 버튼을 누르면 §10의 office.js 없는 순수 웹 패널이 태스크팬으로 열린다.

## 0.12.0 (2026-08-17)

**`xgen-seepage panel` - 리본 버튼 없이도 채팅 패널을 무조건 열 수 있게.**
Excel 리본 버튼(Office 애드인)은 Office 버전에 따라 안 뜰 수 있는데(구형
볼륨 라이선스 등), office.js를 걷어내 패널이 순수 웹 UI가 됐으므로 브라우저로
그냥 열면 그대로 쓸 수 있다. `xgen-seepage panel`(또는 `run --open-panel`)이
패널을 기본 브라우저로 연다. `run`도 시작 시 패널 URL을 명확히 출력.
즉 리본 surfacing은 Office에 달렸지만 패널 사용 자체는 어느 머신이든 보장.

- `panel` 서브커맨드 + `run --open-panel` 추가. 얼린 exe에도 반영 확인.
- README를 현재 상태로 갱신(채팅 패널·에이전트/모델 선택·서버 선택·panel).

## 0.11.1 (2026-08-17)

**코어 검증: 별도 커넥터 프로세스가 "사용자가 연 Excel"을 실제로 편집.**
이전까지의 셀 편집 검증은 커넥터와 Excel이 같은 프로세스였는데, 실제
배포는 사용자가 Excel을 따로 열고 커넥터가 별도로 돈다. 그 형태 그대로
재현해서 검증했다(실제 Excel 머신, 세션 1): 한 프로세스가
`EXCEL.EXE user_doc.xlsx`로 워크북을 열어두고(A1="USER_BEFORE"), 별도
프로세스가 `live_adapter`로 그 워크북을 찾아 A1을 "42"로 편집 →
read-back으로 확인. 즉 사용자의 열린 Excel 문서를 커넥터가 침투해 편집하는
제품의 핵심이 실제 배포 형태로 동작함을 증명. 코드 변경 없음(검증만),
`ARCHITECTURE.md` §11.

## 0.11.0 (2026-08-17)

**실동작 처음으로 끝까지 증명 + 그걸 막던 503 해결 + provider/서버 선택.**
사용자 지적: 매번 "No model loaded 503"이라 실제 동작을 한 번도 못 봤는데
"다 됐다"고 하는 건 문제다. 맞는 지적이었고, 이번에 실제로 풀었다. 자세한
내용은 `ARCHITECTURE.md` §11.

- **503 원인 규명 + 해결**: 워크플로우 harness provider가 비면 서버 기본값
  (안 떠 있는 vLLM)으로 떨어져 죽었다. `execute_stream`이 provider/model을
  받아 워크플로우의 `agents/harness` 노드에 `node_parameter`로 주입하도록
  추가. `provider=anthropic`(서버에 키 있음) 주입하니 바로 동작.
- **에이전트가 진짜 Excel 셀 편집하는 전체 루프를 read-back으로 직접 검증**
  (실제 Excel 머신): A1 "BEFORE"→42, B2 "X"→"DONE" 둘 다 셀 읽어서 확인.
  로그인→브릿지→에이전트→도구 호출→live_adapter→실제 셀 변경까지 실측.
- 패널에 provider/model 드롭다운 추가(`GET /providers`). Playwright로
  anthropic 골라 실행 시 503 없이 셀 편집 확인.
- 로그인 시 서버 선택: `known_servers`(써 본 XGEN 서버 목록)를 번호로
  보여주고 고르게. 고객사 주소를 레포에 하드코딩하지 않고 로컬에만 쌓음.

## 0.10.0 (2026-08-17)

**채팅 패널에서 office.js 의존 완전 제거 - 순수 XGEN 의존 웹 UI로.**
사용자 지적: "Office 비의존이 아니라 XGEN 의존으로 가야 한다", 그리고
패널이 "Office.js 로딩중"에서 멈추는 문제. 맞는 지적이었다 - 이 패널은
Office JavaScript API를 하나도 안 쓰는데(셀 편집은 브릿지가 함) office.js를
넣어둔 게 불필요했고 오히려 `Office.onReady()`가 UI를 멈추게 하고 있었다.

- office.js `<script>`와 `Office.onReady()` 게이팅을 전부 제거. 이제 DOM
  로드 시 `/health`로 커넥터 연결만 확인하고 바로 뜬다. 0.9.0에서 벤더링했던
  `taskpane/vendor/office.js`도 삭제.
- Playwright 실측: 상태가 "XGEN 연결됨"으로 바로 뜨고, 페이지 로드 시
  외부(Microsoft 포함) 요청 0건, 에이전트 드롭다운·채팅 전송 정상.
- 이제 패널까지 포함해 전체 경로가 인터넷·office.js·Microsoft 애드인
  인프라 어디에도 의존하지 않는다. XGEN 커넥터에만 붙어 있으면 동작.

## 0.9.0 (2026-08-17)

**태스크팬에 에이전트 선택 토글 + office.js 로컬 벤더링(폐쇄망 대응).**
사용자 요청: XGEN 로그인만 돼 있으면 패널에서 에이전트를 골라 쓸 수 있게,
그리고 인터넷 의존 없이 동작하게. 자세한 내용은 `ARCHITECTURE.md` §11.

- 서버에 `GET /workflows`(로그인 계정의 워크플로우 목록) 추가, `/chat/stream`이
  요청 바디의 `workflow_id`를 받아 패널에서 고른 에이전트로 실행. CLI로
  워크플로우를 미리 박아둘 필요 없음(`get_chat_workflow_id`도 이제 선택).
- 태스크팬 UI에 에이전트 드롭다운 추가. Playwright로 실측: 실제 계정의
  워크플로우 2개를 불러오고, 선택한 에이전트로 실제 SSE 스트리밍 확인.
- `office.js`를 Microsoft CDN이 아니라 로컬 서버에서 서빙(`taskpane/vendor/
  office.js`). Playwright 네트워크 로그로 페이지 로드 시 microsoft.com 호출
  0건 확인. 단, 완전 에어갭 동작은 실제 폐쇄망 Excel에서 별도 확인 필요.
- **핵심**: 에이전트의 Excel 셀 편집 본체(로그인→브릿지→어댑터)는 office.js도
  인터넷도 필요 없이 이미 동작. office.js는 태스크팬 채팅 UI에만 쓰인다.

## 0.8.0 (2026-08-17)

**Excel 태스크팬 UI 실제 구현 + Playwright로 실기 검증 (Phase 4 대부분).**
`taskpane/`에 manifest.xml, index.html, taskpane.js(SSE 파싱 채팅), taskpane.css,
아이콘을 추가했다. 자세한 내용은 `ARCHITECTURE.md` §11 후반부.

- **Playwright로 실제 브라우저에서 검증**: Office.js 로드 → `Office.onReady()`
  콜백 → 채팅 입력·전송 → 로컬 `/chat/stream` → 실제 dev-xgen.x2bee.com SSE
  스트리밍까지 전부 실제로 동작 확인.
- **남은 미검증**: 매니페스트가 실제 Excel 리본 버튼으로 뜨는 것. 검증에 쓴
  머신 Office가 ProPlus2021 볼륨 라이선스(16.0.14334)라 개발자 레지스트리
  사이드로드를 안 읽고, 공유 폴더 카탈로그용 애드인 삽입 대화상자도 콘텐츠를
  못 그려 막혔다. xgen-seepage 코드 문제가 아니라 그 데스크톱 Office
  클라이언트의 환경 이슈로 규명됨(§11).

## 0.7.0 (2026-08-17)

**`live_adapter`(xlwings/COM)를 실제 Microsoft Excel로 처음 끝까지
검증.** 실제 Excel이 설치된 사용자의 다른 Windows PC에 접속해 24개 도구 중
`live_*` 8개 전부를 실제 열린 통합문서 상대로 왕복 확인했다. 자세한
내용은 `ARCHITECTURE.md` §11.

- **실제로 찾은 버그 3개**: xlwings 0.36.16엔 `book.saved`가 없어(`.api.Saved`로
  수정) `list_open_workbooks`가 즉사했고, 다중 셀 range의 `.formula`가
  list가 아니라 tuple로 와서 `read_range`가 기형적으로 중첩된 배열을
  냈으며, Excel의 모든 숫자가 float으로 오는 특성 때문에 정수값이
  "3.0"으로 보였다(`_cellfmt.cell_text` 수정).
- **별개로 찾은 인프라 버그**: `ConnectorMcpBridge.stop()`이 진행 중인 WS
  연결을 안 닫아서 정상 종료(`xgen-seepage run`을 Ctrl+C)가 영원히 안
  끝나는 hang이 있었다. 실제 정상 종료 시나리오를 테스트하다 발견, 수정 후
  재확인.
- 43개 테스트 통과, ruff/mypy 클린.

## 0.6.1 (2026-08-17)

**심각한 패키징 버그 수정 - 얼린 exe가 자기 자신을 실행조차 못 했다.**
태스크팬 기능을 실제로 얼려서 처음 검증하며 `xgen-seepage-connector.exe
--help`를 직접 돌려보니 `ModuleNotFoundError: No module named
'xgen_seepage.connector'`로 즉시 죽었다. 원인은 스펙 파일의 `pathex=[]` -
`packaging/`에서 빌드하면 프로젝트 루트(`xgen_seepage/`가 있는 곳)가 빌드
시점 파이썬의 `sys.path`에 없어서 생기는 문제였다. `_uno_worker.py`
버그처럼 빌드는 항상 성공하고 실행하면 그 자리에서 죽는 부류라 빌드
로그만으로는 못 잡는다. `SPECPATH` 기준 절대경로로 고정해 빌드 실행
위치와 무관하게 만들었다. 지금까지의 얼린-exe 검증은 전부 자체 pathex를
따로 지정한 별도 probe exe로 했지 실제 exe의 CLI를 직접 돌려본 적이
없어서 이 문제를 여태 못 잡았다는 것도 정직하게 기록해둔다.

고친 뒤 실제로 얼려서 `--help`/`status`/`login`(HTTP는 성공, OS 키체인
저장은 기존에 알려진 비대화형 세션 제약으로만 실패)과, 별도 probe로 얼린
프로세스 안에서 `TaskpaneServer`를 띄워 `/health`/정적 파일 서빙/
`/chat/stream`(실제 dev-xgen.x2bee.com에서 SSE 실시간 수신)까지 전부
확인했다. 우려했던 uvicorn 런타임 디스패치 문제는 `collect_all` +
명시적 hidden-imports로 미리 막아둔 덕에 재현되지 않았다. 자세한 내용은
`ARCHITECTURE.md` §7.

## 0.6.0 (2026-08-17)

**설계 정정: "워크플로우 자동생성"이 아니라 "연결"이다.** 0.5.0에서 계획한
Phase 3(채팅용 워크플로우가 없으면 xgen-seepage가 자동 생성)은 틀린
방향이었다 - 사용자가 XGEN에 이미 갖고 있는 워크플로우 중 하나에 태스크팬을
연결해야 한다. 자세한 내용은 `ARCHITECTURE.md` §10.

- **서버 로그로 ERROR401의 진짜 원인을 확정했다**(SSH로 dev-177 파드 로그
  직접 확인): (1) 방금 만든 워크플로우는 dev 서버 전반에 걸쳐 자기 자신도
  못 찾는 실제 데이터 일관성 버그가 있고(내 계정만의 문제가 아님 - 기존
  스케줄 작업도 같은 이유로 죽고 있었다), (2) 그 에러를 사용자에게
  보여주려는 `error_message_replacer`가 서버 쪽에서 `None`으로 바인딩돼
  있어 `TypeError`로 한 번 더 죽으면서 진짜 원인이 감춰지고 있었다. 둘 다
  xgen-seepage 밖의 서버 쪽 문제.
- **이미 존재가 확인된 워크플로우(`shinhan_blue_agent_v1`)로 재시도하니
  완전히 정상 동작했다** - 실행 로그 전체가 실시간으로 프록시를 거쳐
  도착했고 스트림이 끝까지 흘렀다(§8에서 이미 알려진 vLLM 모델 미기동
  이슈만 남음, 새 문제 아님). "연결" 설계는 이 버그를 구조적으로 비켜간다.
- `agentflow_client.py::list_workflows()` 신규, CLI
  `xgen-seepage chat-workflow list`/`set` 추가. `config.py`의 이제 안 쓰는
  `chat_provider`/`chat_model` 필드(자동생성 템플릿용) 제거.

## 0.5.0 (2026-08-17)

**Excel 태스크팬 채팅 Phase 2/5: SSE 프록시.** `connector/agentflow_client.py`
신규(하네스 워크플로우 CRUD + 실행 스트림)와 `taskpane_server.py`의
`/chat/stream`(XGEN의 SSE를 원시 바이트로 그대로 릴레이)을 실제
dev-xgen.x2bee.com으로 검증. 자세한 내용은 `ARCHITECTURE.md` §10.

- **실제로 찾은 버그**: `_ChatStreamHandler`를 `__call__`을 가진 클래스
  인스턴스로 라우트에 등록했더니 Starlette가 일반 요청 핸들러가 아니라
  raw ASGI 앱으로 오인해서 즉시 500 (`takes 2 positional arguments but 4
  were given`). 바운드 메서드(`chat_handler.handle`)로 등록하도록 수정.
- **증명된 것**: 로컬 `/chat/stream`에 POST하면 실제 XGEN 서버까지 갔다가
  SSE 응답이 바이트 단위로 그대로 돌아온다(왜곡 없음) - 프록시 계층 자체는
  완전히 동작.
- **증명 안 된 것**: 실제 에이전트 응답까지 오는 전체 실행은 지금
  dev-xgen.x2bee.com에서 `ERROR401`(스트리밍 실행 중 예외 catch-all, 백엔드
  소스로 확인)로 계속 실패한다. 원인이 서버 쪽 일시적 불안정인지 다른 문제인지
  클라이언트에서는 구분 불가 - 다음 세션 최우선 확인 항목.
- `HttpClient.stream()` 추가(기존 `json()`/`get()`/`post()`는 그대로).
- 43개 테스트 통과, ruff/mypy 클린(기존 bridge.py 미해결 오류 2개 제외).

## 0.4.0 (2026-08-17)

**Excel 태스크팬 채팅 기능 착수 (Phase 1/5: 로컬 HTTPS 서버 골격).** 채팅을
XGEN 웹 UI가 아니라 Excel 안에 도킹된 실제 패널에서 하고 싶다는 요청에 따라,
Office.js 태스크팬 애드인 경로를 다시 꺼냈다(§1-§2에서 설치 마찰 때문에
일부러 뺐던 경로). 자세한 아키텍처/진행 상황은 `ARCHITECTURE.md` §10.

- `connector/certs.py` 신규: 자체서명 인증서를 `cryptography`로 순수
  파이썬 생성(새 pip 설치 없음 - `mcp`의 전이 의존성으로 이미 있던 걸
  직접 의존성으로 승격), Windows `certutil -user -addstore Root`로 신뢰
  등록 시도. **실측으로 찾은 것**: 이 비대화형 세션에서는 Root 저장소
  추가만 `ERROR_NOT_SUPPORTED`로 실패한다(CA/My 저장소는 같은 세션에서
  성공 - 파일 포맷 문제 아님을 확인). OS 키체인과 같은 부류의 "대화형
  로그온 세션 필요" 제약으로 보이며, 정상 데스크톱 세션에서의 동작은
  별도 검증 필요.
- `connector/taskpane_server.py` 신규: Starlette+uvicorn 로컬 HTTPS 서버.
  `cmd_run`이 기존 WS 브릿지와 `asyncio.gather`로 나란히 띄운다(`--no-taskpane`
  로 끌 수 있음).
- 실제 dev-xgen.x2bee.com에 로그인한 채로 브릿지(`catalog_synced=True,
  server_tool_count=24`)와 로컬 `/health`가 동시에 정상 동작하는 것을
  확인. 43개 테스트 통과, ruff/mypy 클린.

## 0.3.1 (2026-08-16)

**실제로 얼려서(PyInstaller) 검증하다 찾은 진짜 버그 수정.** LibreOffice
백엔드가 서브프로세스로 띄우는 `_uno_worker.py`는 import가 아니라 파일
경로로만 참조돼서, PyInstaller가 얼린 exe 안에 그 파일을 넣지 않았다 -
폐쇄망 배포용으로 만든 기능이 정작 폐쇄망 배포 산출물 안에서는 동작하지
않는 상태였다. `packaging/xgen-seepage-connector.spec`에 그 파일을 `datas`로
명시해 고쳤고, 고친 뒤 얼린 exe로 실제 LibreOffice를 열어 셀 쓰기→읽기→저장
왕복까지 확인했다(`ARCHITECTURE.md` §7).

README에서 내부 참고 레포/다른 사내 클라이언트 언급을 정리하고 도구 목록을
현재 상태(24개)로 맞췄다.

## 0.3.0 (2026-08-14)

**LibreOffice(UNO) 백엔드 추가. Microsoft Excel이 없는 환경(폐쇄망 등)에서도
같은 셀 IO 경험을 제공.** 이 저장소를 만든 머신엔 Excel이 없다(3가지 방법으로
확인). 대신 이미 설치돼 있던 LibreOffice로 실제 동작하는 대안 백엔드를 새로
만들었다(자세한 아키텍처는 `ARCHITECTURE.md` §9).

- `libreoffice_adapter.py` / `_uno_worker.py` 신규: UNO로 xlsx를 LibreOffice에서
  직접 열어 스키마 조회/셀 읽기·쓰기/범위 읽기·쓰기/행 추가/저장을 제공.
  `uno` 모듈이 LibreOffice 번들 파이썬 전용이라 상주 서브프로세스(stdin/stdout
  JSON 프로토콜)로 격리했다.
- `tools.py`에 `open_libreoffice_document`/`get_libreoffice_schema`/
  `set_libreoffice_cell`/`read_libreoffice_range`/`write_libreoffice_range`/
  `append_libreoffice_row`/`save_libreoffice_document` 등 10개 도구를 별도
  tool 군으로 등록(총 24개 도구).
- 실제 파일(제주은행 결함관리대장 사본) 왕복 편집으로 수동 검증. 기존
  43개 단위/통합 테스트 통과 유지, ruff/mypy 클린.

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

- 사설/자체서명 CA를 쓰는 폐쇄망 XGEN 서버 접속을 위한 `allow_private_certificate`
  를 실제로 구현했다(전에는 설정 필드만 있고 어디에도 안 이어져 있었다).
  `connection_security.py` 신규(사설 CA를 명시적으로 신뢰하는 TLS 정책).
  `login --allow-private-certificate` 플래그 +
  `XGEN_SEEPAGE_ALLOW_PRIVATE_CERTIFICATE` 환경변수(배포 기본값용) 추가. wss://에만
  적용하고 ws://에는 안 넣는 로직까지 실제 테스트로 검증
  (`tests/test_connector_tls.py`).
- **실제 크래시 버그 발견 및 전면 수정**: `argparse` 도움말 문자열에 든
  em dash(U+2014)가 Windows 한글 콘솔(cp949) 출력 시 `UnicodeEncodeError`로
  CLI 자체를 죽였다(`xgen-seepage login --help` 재현). 이 프로젝트 전체
  (코드 주석·문서·커밋 예정 텍스트 포함)에서 em dash를 전부 제거했다
  (사용자 지침 `no-em-dash` 위반이기도 했음. 42/42 테스트 통과 유지 확인).

## 0.2.0 (2026-08-13)

방향 전환: 로컬 MCP에 얹는 방식에서 **독립 실행형 커넥터**로 재설계(사용자
정정. 폐쇄망 단일 설치 프로그램이어야 하고, 외부 클라이언트가 런타임
의존성이면 안 됨).

- `connector/` 신규: XGEN 서버에 직접 로그인(`hash.py`/`auth.py`,
  `/api/auth/login` 등 XGEN 프로토콜을 독립 재구현) +
  `/api/tools/ws/connector-mcp/{user_id}` WebSocket 브릿지
  (`bridge.py`, hello/ready/mcp_call/mcp_result) + 로컬 설정/OS 키체인 토큰
  저장(`config.py`) + CLI(`app.py`: `login`/`run`/`status`/`logout`).
- 새 콘솔 스크립트 `xgen-seepage`(기본 경로). 기존 `xgen-seepage-mcp`는
  로컬 MCP 클라이언트/Claude Desktop에 얹고 싶을 때 쓰는 대안으로 유지.
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
  (`xgen-seepage-mcp`). 로컬 MCP를 지원하는 클라이언트에 그대로 등록 가능.
- 리서치 기반 아키텍처 문서 `ARCHITECTURE.md`.
