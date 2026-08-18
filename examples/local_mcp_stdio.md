# 대안 경로. stdio MCP 서버로 다른 MCP 클라이언트에 등록하기

> ⚠️ **이건 대안 경로다, 기본 경로가 아니다.** xgen-seepage의 기본 사용법은
> `xgen-seepage login` + `xgen-seepage run`으로 XGEN 서버에 **직접** 붙는
> 것이다(README 참조). 이 문서는 **로컬/stdio MCP 서버 등록을 지원하는 다른
> 클라이언트를 이미 쓰고 있어서**, 그 앱의 채팅 창에서 곧바로 xgen-seepage
> 도구를 쓰고 싶은 경우를 위한 것이다.

`xgen-seepage-mcp`는 표준 stdio MCP 서버다. 로컬 MCP를 지원하는 클라이언트라면
이 프로세스를 등록만 하면 별도 워크플로우 편집이나 코드 수정 없이 도구를 쓸 수
있다.

## 1. xgen-seepage 설치

```bash
pip install xgen-seepage[live]   # 라이브(열려 있는 통합문서) 편집까지 쓰려면
# 또는
pip install xgen-seepage         # CSV 도구만 필요하면
```

Windows/macOS에서 로컬 Excel이 설치돼 있어야 라이브 편집(`get_live_*`/`set_live_*`/
`*_live_range`)이 동작한다. 인터넷은 필요 없다.

## 2. 실행 경로 확인

```bash
python -m xgen_seepage.mcp_server
# 정상이면 stdio로 대기 상태가 된다(Ctrl+C로 종료)
```

또는 콘솔 스크립트:

```bash
xgen-seepage-mcp
```

## 3. MCP 클라이언트 설정

로컬(stdio) MCP 서버 추가 화면에서:

| 필드 | 값 |
|---|---|
| 이름 | `xgen-seepage` |
| 전송 | stdio |
| command | `python` (또는 venv의 절대경로) |
| args | `-m xgen_seepage.mcp_server` |

저장하면 클라이언트가 이 프로세스를 스폰해 도구 카탈로그를 수집하고, 이후
로그인한 세션에서 Excel 라이브/LibreOffice/CSV 도구가 자동으로 주어진다.

## 4. 실제 사용 흐름 예시

```
사용자: (Excel에서 매출_2026.xlsx를 열어 둔 상태) "B열에 세율 10% 적용한 열 하나 추가해줘"

에이전트 도구 호출 순서:
  1. list_open_workbooks()               → workbook_id 확보
  2. get_live_schema(workbook_id, 0)      → 레이아웃 파악(B열이 "매출"인지 확인)
  3. read_live_range(...)                 → 기존 B열 값 읽기
  4. write_live_range(..., rows=[["=B2*1.1"], ...])  → C열에 수식 채우기
  5. activate_live_cell(...)              → 사용자가 결과를 바로 눈으로 확인
```

CSV라면:

```
사용자: "customers.csv 열어서 3번째 컬럼에 등급 매겨줘"

  1. inspect_csv(path)                    → 헤더/미리보기 확인
  2. set_csv_cell(...) 반복 또는 write_csv_table(...)로 통째 갱신
```

CSV에 수식이나 서식이 필요하면 `open_csv_in_excel(path)`로 먼저 Excel에 올린
뒤, 반환된 `workbook_id`로 라이브 도구를 이어서 쓴다.
