"""MCP tool 정의 + dispatcher.

스키마 형태(각 tool = {name, description, input_schema} dict, dispatcher는
이름→dict 반환 함수, 예외는 {"error":..., "message":...}로 직렬화)는 저장소
전체에서 일관되게 유지한다. 같은 XGEN 세션에 여러 도구를 같이 등록해도
에이전트가 서버마다 다른 스타일로 다룰 필요가 없게 하기 위함.

라이브(Excel) 도구는 성격이 뚜렷한 것만 단일 도구로 두고, 파라미터가 균질한
조작(서식·행열 삽입삭제·숨김·크기·시트·이름)은 하나의 액션 도구로 묶는다.
"""
from __future__ import annotations

from typing import Any, Callable

from . import csv_adapter, libreoffice_adapter, live_adapter
from .base import SeepageError

# -------- 라이브(열려 있는 통합문서) 도구 --------
# 조회·값편집 등 단일 목적 도구


def list_open_workbooks() -> dict[str, Any]:
    books = live_adapter.list_open_workbooks()
    return {"workbooks": [b.to_dict() for b in books]}


def get_workbook_overview(workbook_id: str | None = None, preview_rows: int = 6) -> dict[str, Any]:
    return live_adapter.get_workbook_overview(workbook_id, preview_rows)


def get_live_schema(workbook_id: str | None = None, sheet: int | str = 0,
                     preview_rows: int = 12, max_cell_len: int = 60) -> dict[str, Any]:
    schema = live_adapter.get_sheet_schema(
        workbook_id, sheet, preview_rows=preview_rows, max_cell_len=max_cell_len)
    return schema.to_dict()


def get_live_cell(workbook_id: str | None, sheet: int | str, row: int, col: int) -> dict[str, Any]:
    return live_adapter.get_cell(workbook_id, sheet, row, col).to_dict()


def get_live_table_region(sheet: int | str, row: int, col: int,
                          workbook_id: str | None = None) -> dict[str, Any]:
    return live_adapter.get_table_region(workbook_id, sheet, row, col)


def set_live_cell(workbook_id: str | None, sheet: int | str, row: int, col: int,
                   value: str, as_formula: bool = False,
                   allow_merge_redirect: bool = False) -> dict[str, Any]:
    old = live_adapter.set_cell(workbook_id, sheet, row, col, value,
                                 as_formula=as_formula,
                                 allow_merge_redirect=allow_merge_redirect)
    return {"old_value": old}


def read_live_range(workbook_id: str | None, sheet: int | str,
                     row0: int, col0: int, row1: int, col1: int) -> dict[str, Any]:
    return live_adapter.read_range(workbook_id, sheet, row0, col0, row1, col1)


def write_live_range(workbook_id: str | None, sheet: int | str,
                      row0: int, col0: int, rows: list[list[str]]) -> dict[str, Any]:
    live_adapter.write_range(workbook_id, sheet, row0, col0, rows)
    return {"written_rows": len(rows)}


def append_live_row(workbook_id: str | None, sheet: int | str, values: list[str]) -> dict[str, Any]:
    row = live_adapter.append_row(workbook_id, sheet, values)
    return {"row": row}


def find_replace_live(sheet: int | str, find: str, replace: str, match_case: bool = False,
                      workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.find_replace(workbook_id, sheet, find, replace, match_case)
    return {"ok": True}


def recalculate_live(workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.recalculate(workbook_id)
    return {"ok": True}


def activate_live_cell(workbook_id: str | None = None, sheet: int | str | None = None,
                        row: int | None = None, col: int | None = None) -> dict[str, Any]:
    live_adapter.activate(workbook_id, sheet, row, col)
    return {"ok": True}


def color_live_rows_where(sheet: int | str, condition_col: int, match_value: str, color: Any,
                          target_col0: int, target_col1: int, data_start_row: int = 0,
                          contains: bool = False, clear_non_matching: bool = True,
                          workbook_id: str | None = None) -> dict[str, Any]:
    return live_adapter.color_rows_where(
        workbook_id, sheet, condition_col, match_value, color,
        target_col0, target_col1, data_start_row, contains, clear_non_matching)


def sort_live_range(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                    key_col: int, ascending: bool = True, has_header: bool = True,
                    workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.sort_range(workbook_id, sheet, row0, col0, row1, col1, key_col, ascending, has_header)
    return {"ok": True}


def copy_live_range(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                    dest_row: int, dest_col: int, dest_sheet: int | str | None = None,
                    workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.copy_range(workbook_id, sheet, row0, col0, row1, col1,
                            dest_row, dest_col, dest_sheet)
    return {"ok": True}


def autofit_live(sheet: int | str, row0: int | None = None, col0: int | None = None,
                 row1: int | None = None, col1: int | None = None,
                 workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.autofit(workbook_id, sheet, row0, col0, row1, col1)
    return {"ok": True}


def freeze_live_panes(sheet: int | str, rows: int = 1, cols: int = 0,
                      workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.freeze_panes(workbook_id, sheet, rows, cols)
    return {"ok": True}


def set_live_autofilter(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                        on: bool = True, workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.set_autofilter(workbook_id, sheet, row0, col0, row1, col1, on)
    return {"ok": True}


# 응집 도구(dispatcher): 파라미터가 균질한 조작을 하나의 action 도구로 묶는다.


def format_live_range(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                      action: str, color: Any = None, bold: bool | None = None,
                      font_color: Any = None, format_code: str | None = None,
                      which: str = "all", style: str = "thin",
                      wrap: bool = True, autofit_rows: bool = True,
                      values: list[str] | None = None,
                      contents: bool = True, formats: bool = False,
                      workbook_id: str | None = None) -> dict[str, Any]:
    if action == "fill_color":
        live_adapter.set_fill_color(workbook_id, sheet, row0, col0, row1, col1, color)
    elif action == "font":
        live_adapter.format_range(workbook_id, sheet, row0, col0, row1, col1,
                                  fill_color=color, bold=bold, font_color=font_color)
    elif action == "number_format":
        live_adapter.set_number_format(workbook_id, sheet, row0, col0, row1, col1, format_code)
    elif action == "borders":
        live_adapter.set_borders(workbook_id, sheet, row0, col0, row1, col1, which, style)
    elif action == "wrap":
        live_adapter.set_wrap_text(workbook_id, sheet, row0, col0, row1, col1, wrap, autofit_rows)
    elif action == "merge":
        live_adapter.merge_cells(workbook_id, sheet, row0, col0, row1, col1)
    elif action == "unmerge":
        live_adapter.unmerge_cells(workbook_id, sheet, row0, col0, row1, col1)
    elif action == "data_validation":
        live_adapter.set_data_validation_list(workbook_id, sheet, row0, col0, row1, col1, values or [])
    elif action == "clear":
        live_adapter.clear_range(workbook_id, sheet, row0, col0, row1, col1, contents, formats)
    else:
        raise ValueError(f"format_live_range: 알 수 없는 action {action!r}")
    return {"ok": True, "action": action}


def edit_live_rows_columns(sheet: int | str, action: str, index: int, count: int = 1,
                           workbook_id: str | None = None) -> dict[str, Any]:
    if action == "insert_rows":
        live_adapter.insert_rows(workbook_id, sheet, index, count)
    elif action == "delete_rows":
        live_adapter.delete_rows(workbook_id, sheet, index, count)
    elif action == "insert_columns":
        live_adapter.insert_columns(workbook_id, sheet, index, count)
    elif action == "delete_columns":
        live_adapter.delete_columns(workbook_id, sheet, index, count)
    else:
        raise ValueError(f"edit_live_rows_columns: 알 수 없는 action {action!r}")
    return {"ok": True, "action": action}


def set_live_visibility(sheet: int | str, action: str, start: int, end: int, visible: bool,
                        workbook_id: str | None = None) -> dict[str, Any]:
    if action == "rows":
        live_adapter.set_rows_visible(workbook_id, sheet, start, end, visible)
    elif action == "columns":
        live_adapter.set_columns_visible(workbook_id, sheet, start, end, visible)
    else:
        raise ValueError(f"set_live_visibility: 알 수 없는 action {action!r}")
    return {"ok": True, "action": action}


def resize_live(sheet: int | str, action: str, index: int, size: float,
                workbook_id: str | None = None) -> dict[str, Any]:
    if action == "column_width":
        live_adapter.set_column_width(workbook_id, sheet, index, size)
    elif action == "row_height":
        live_adapter.set_row_height(workbook_id, sheet, index, size)
    else:
        raise ValueError(f"resize_live: 알 수 없는 action {action!r}")
    return {"ok": True, "action": action}


def manage_live_sheet(action: str, sheet: int | str | None = None, name: str | None = None,
                      new_name: str | None = None, before: int | str | None = None,
                      after: int | str | None = None,
                      workbook_id: str | None = None) -> dict[str, Any]:
    if action == "list":
        return {"sheets": live_adapter.list_sheets(workbook_id)}
    if action == "add":
        return {"name": live_adapter.add_sheet(workbook_id, name, before, after)}
    if action == "rename":
        return {"name": live_adapter.rename_sheet(workbook_id, sheet, new_name)}
    if action == "delete":
        live_adapter.delete_sheet(workbook_id, sheet)
        return {"ok": True}
    if action == "move":
        live_adapter.move_sheet(workbook_id, sheet, before, after)
        return {"ok": True}
    raise ValueError(f"manage_live_sheet: 알 수 없는 action {action!r}")


def manage_live_name(action: str, name: str | None = None, refers_to: str | None = None,
                     workbook_id: str | None = None) -> dict[str, Any]:
    if action == "list":
        return {"names": live_adapter.list_names(workbook_id)}
    if action == "define":
        live_adapter.define_name(workbook_id, name, refers_to)
        return {"ok": True}
    if action == "delete":
        live_adapter.delete_name(workbook_id, name)
        return {"ok": True}
    raise ValueError(f"manage_live_name: 알 수 없는 action {action!r}")


# -------- CSV(파일) 도구 --------


def inspect_csv(path: str, preview_rows: int = 12, max_cell_len: int = 60,
                 has_header: bool = True) -> dict[str, Any]:
    return csv_adapter.get_schema(path, preview_rows=preview_rows, max_cell_len=max_cell_len,
                                   has_header=has_header).to_dict()


def get_csv_cell(path: str, row: int, col: int, has_header: bool = True) -> dict[str, Any]:
    return csv_adapter.get_cell(path, row, col, has_header=has_header).to_dict()


def set_csv_cell(path: str, row: int, col: int, value: str, has_header: bool = True) -> dict[str, Any]:
    old = csv_adapter.set_cell(path, row, col, value, has_header=has_header)
    return {"old_value": old}


def append_csv_row(path: str, values: list[str], has_header: bool = True) -> dict[str, Any]:
    row = csv_adapter.append_row(path, values, has_header=has_header)
    return {"row": row}


def write_csv_table(path: str, rows: list[list[str]], header: list[str] | None = None,
                     encoding: str = "utf-8-sig") -> dict[str, Any]:
    csv_adapter.write_table(path, rows, header=header, encoding=encoding)
    return {"path": path, "rows": len(rows)}


def open_csv_in_excel(path: str) -> dict[str, Any]:
    return {"workbook": csv_adapter.open_in_excel(path)}


# -------- LibreOffice(파일 직접 침투) 도구 --------
#
# live_adapter(위)는 "이미 열려 있는" Excel 통합문서에 붙는다 - Excel이 설치돼
# 있어야 한다. 이 섹션은 Microsoft Excel이 아예 없는 환경(폐쇄망 등)에서도
# xlsx 파일 자체를 LibreOffice로 열어 같은 셀 입출력 경험을 제공한다. doc는
# open_libreoffice_document가 돌려준 path를 그대로 넘긴다.


def open_libreoffice_document(path: str, hidden: bool = True) -> dict[str, Any]:
    return libreoffice_adapter.open_document(path, hidden=hidden)


def close_libreoffice_document(doc: str | None = None, save: bool = False) -> dict[str, Any]:
    libreoffice_adapter.close_document(doc, save=save)
    return {"ok": True}


def list_open_libreoffice_documents() -> dict[str, Any]:
    return {"documents": libreoffice_adapter.list_open_documents()}


def get_libreoffice_schema(doc: str | None = None, sheet: int | str = 0,
                            preview_rows: int = 12, max_cell_len: int = 60) -> dict[str, Any]:
    schema = libreoffice_adapter.get_sheet_schema(
        doc, sheet, preview_rows=preview_rows, max_cell_len=max_cell_len)
    return schema.to_dict()


def get_libreoffice_cell(doc: str | None, sheet: int | str, row: int, col: int) -> dict[str, Any]:
    return libreoffice_adapter.get_cell(doc, sheet, row, col).to_dict()


def set_libreoffice_cell(doc: str | None, sheet: int | str, row: int, col: int,
                          value: str, as_formula: bool = False,
                          allow_merge_redirect: bool = False) -> dict[str, Any]:
    old = libreoffice_adapter.set_cell(doc, sheet, row, col, value,
                                        as_formula=as_formula,
                                        allow_merge_redirect=allow_merge_redirect)
    return {"old_value": old}


def read_libreoffice_range(doc: str | None, sheet: int | str,
                            row0: int, col0: int, row1: int, col1: int) -> dict[str, Any]:
    return libreoffice_adapter.read_range(doc, sheet, row0, col0, row1, col1)


def write_libreoffice_range(doc: str | None, sheet: int | str,
                             row0: int, col0: int, rows: list[list[str]]) -> dict[str, Any]:
    libreoffice_adapter.write_range(doc, sheet, row0, col0, rows)
    return {"written_rows": len(rows)}


def append_libreoffice_row(doc: str | None, sheet: int | str, values: list[str]) -> dict[str, Any]:
    row = libreoffice_adapter.append_row(doc, sheet, values)
    return {"row": row}


def save_libreoffice_document(doc: str | None = None) -> dict[str, Any]:
    libreoffice_adapter.save(doc)
    return {"ok": True}


# -------- 도구 정의 --------

TOOL_DEFINITIONS: list[dict[str, Any]] = [
    {
        "name": "list_open_workbooks",
        "description": (
            "로컬에서 실행 중인 모든 Excel 인스턴스의 열려 있는 통합문서 목록을 "
            "반환한다. 라이브 편집 도구(get_live_schema/set_live_cell/...)를 쓰기 "
            "전에 항상 먼저 호출해서 workbook_id를 확보해야 한다. workbook_id를 "
            "생략하면 이후 호출은 '현재 활성 통합문서'를 임의로 집는다. "
            "★편집 원칙: 사용자의 원본을 보존한다. 먼저 get_live_schema/read_live_range로 "
            "기존 내용·수식을 파악하고, 요청한 부분만 정확히 바꿔라. 원본을 지우거나 "
            "통째로 다시 쓰지 말고, 시트 삭제/범위 전체 덮어쓰기는 사용자가 명시적으로 "
            "요청할 때만 한다."
        ),
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_workbook_overview",
        "description": (
            "통합문서 **전체 구조**를 한 번에 파악한다 - 모든 시트의 크기·미리보기·"
            "수식 존재 여부·병합. 시트가 여러 개인 복잡한 파일을 편집하기 전에 먼저 "
            "이걸로 전체를 읽어 파악하고, 그 다음 필요한 시트/범위를 read_live_range로 "
            "자세히 본다. 참고: 크기(rows/cols)는 테두리·서식이 있는 빈 셀까지 포함하는 "
            "used_range 기준이라, 테두리로 그린 표 범위도 반영된다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "preview_rows": {"type": "integer", "description": "시트별 미리보기 행 수(기본 6)"},
            },
        },
    },
    {
        "name": "get_live_schema",
        "description": (
            "지금 열려 있는 통합문서의 시트 하나를 조사한다: 크기, 미리보기 표, "
            "병합 셀, 수식 존재 여부. 셀을 편집하기 전에 먼저 호출해 레이아웃을 "
            "파악해야 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string", "description": "list_open_workbooks가 준 id. 생략하면 활성 통합문서."},
                "sheet": {"description": "시트 인덱스(0-based) 또는 이름", "default": 0},
                "preview_rows": {"type": "integer", "default": 12},
                "max_cell_len": {"type": "integer", "default": 60},
            },
        },
    },
    {
        "name": "get_live_cell",
        "description": "셀 하나의 값·수식·병합 정보를 절단 없이 전체로 반환한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row": {"type": "integer", "description": "0-based 행"},
                "col": {"type": "integer", "description": "0-based 열"},
            },
            "required": ["sheet", "row", "col"],
        },
    },
    {
        "name": "read_live_range",
        "description": (
            "직사각형 범위를 값/수식 2차원 배열로 한 번에 읽는다. 셀을 하나씩 "
            "read하는 것보다 훨씬 빠르다. 여러 셀을 볼 땐 이 도구를 쓴다. "
            "최대 50,000셀."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "row1": {"type": "integer"}, "col1": {"type": "integer"},
            },
            "required": ["sheet", "row0", "col0", "row1", "col1"],
        },
    },
    {
        "name": "get_live_table_region",
        "description": "지정한 셀이 속한 '표'의 범위를 반환한다. 값·테두리·서식으로 그려 놓은 표의 실제 경계(빈 행·열로 둘러싸인 연속 블록, Excel의 현재 영역/Ctrl+A)를 잡아 준다. 표가 어디부터 어디까지인지 정확히 알아야 할 때(헤더 위치·데이터 범위 판단) 먼저 호출한다. 반환: row0/col0/row1/col1(0-based, 포함) + n_rows/n_cols.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row": {"type": "integer", "description": "표 안 아무 셀의 행(0-based)"},
            "col": {"type": "integer", "description": "표 안 아무 셀의 열(0-based)"}},
            "required": ["sheet", "row", "col"]}},
    {
        "name": "set_live_cell",
        "description": (
            "지금 화면에 떠 있는 시트의 셀 하나를 즉시 갱신한다. 파일 저장 없이 "
            "사용자가 바로 화면에서 변경을 본다. 숫자처럼 보이는 문자열은 자동으로 "
            "숫자로 기록해 SUM 등 기존 수식이 계속 동작하게 한다. as_formula=true면 "
            "value를 '='로 시작하는 수식으로 그대로 쓴다. 병합 셀의 non-anchor "
            "좌표에 쓰면 MergedCellWriteError. allow_merge_redirect=true로 anchor에 "
            "리다이렉트하거나 anchor 좌표로 직접 호출한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row": {"type": "integer"},
                "col": {"type": "integer"},
                "value": {"type": "string"},
                "as_formula": {"type": "boolean", "default": False},
                "allow_merge_redirect": {"type": "boolean", "default": False},
            },
            "required": ["sheet", "row", "col", "value"],
        },
    },
    {
        "name": "write_live_range",
        "description": (
            "직사각형 범위를 한 번에 덮어쓴다. ⚠️ 지정한 범위의 기존 값·수식·서식이 "
            "사라진다. 원본을 보존하며 편집하는 게 원칙이니: (1) 실제로 바꿀 셀만 좁게 "
            "잡아라, (2) 이미 채워진 표/영역을 통째로 다시 쓰지 마라(수식·서식이 날아간다), "
            "(3) 몇 개 셀만 바꾸는 거면 set_live_cell을 써라, (4) 주로 '빈 새 영역'을 채울 "
            "때 쓴다. 값 앞에 '='를 붙이면 수식. 최대 50,000셀."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                    "description": "행 우선 2차원 배열",
                },
            },
            "required": ["sheet", "row0", "col0", "rows"],
        },
    },
    {
        "name": "append_live_row",
        "description": "지금 사용 범위 바로 아래에 새 행을 추가한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "values": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["sheet", "values"],
        },
    },
    {
        "name": "find_replace_live",
        "description": "시트의 사용 범위에서 find를 replace로 모두 바꾼다.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "find": {"type": "string"}, "replace": {"type": "string"},
            "match_case": {"type": "boolean", "default": False}},
            "required": ["sheet", "find", "replace"]}},
    {
        "name": "recalculate_live",
        "description": "모든 수식을 강제로 재계산한다(값이 안 갱신되거나 자동계산이 꺼진 것 같을 때).",
        "input_schema": {"type": "object", "properties": {"workbook_id": {"type": "string"}}}},
    {
        "name": "activate_live_cell",
        "description": (
            "Excel 창을 앞으로 가져오고 지정한 시트/셀을 선택한다. 기능적으로는 "
            "필요 없지만 에이전트가 지금 어느 셀을 다루는지 사용자가 눈으로 "
            "따라올 수 있게 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row": {"type": "integer"}, "col": {"type": "integer"},
            },
        },
    },
    {
        "name": "color_live_rows_where",
        "description": (
            "조건 열의 값이 특정 값인 '행들'을 한 번에 색칠한다. 조건부 강조·플래그의 정석 도구. "
            "행 번호를 직접 계산하지 않으므로 헤더가 여러 줄인 실무 표에서도 어긋나지 않는다. "
            "동작: 데이터 행(data_start_row부터)을 훑어 condition_col 값이 match_value와 같은(또는 "
            "contains=true면 포함하는) 행의 대상 열 범위(target_col0..target_col1)를 color로 칠한다. "
            "clear_non_matching=true(기본)면 대상 열의 기존 채우기를 먼저 지우고 매칭 행만 칠해, "
            "잘못 칠해진 것도 함께 정리된다. 열 인덱스는 모두 0-based. "
            "먼저 get_workbook_overview/get_live_schema로 헤더 행 위치와 상태 열 위치를 확인해 "
            "data_start_row와 condition_col을 정확히 잡아라."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "condition_col": {"type": "integer", "description": "조건 열 0-based 인덱스(예: 상태 열)"},
                "match_value": {"type": "string", "description": "이 값과 같은(또는 contains면 포함하는) 행을 칠함"},
                "color": {"description": "'#FFFF00' / 'yellow' / '노란색' 등"},
                "target_col0": {"type": "integer", "description": "칠할 열 범위 시작(0-based)"},
                "target_col1": {"type": "integer", "description": "칠할 열 범위 끝(0-based). 한 열이면 target_col0와 동일"},
                "data_start_row": {"type": "integer", "description": "첫 데이터 행 0-based 인덱스(헤더 아래). 기본 0"},
                "contains": {"type": "boolean", "description": "true면 부분 일치. 기본 false(정확히 일치)"},
                "clear_non_matching": {"type": "boolean", "description": "true(기본)면 대상 열 기존 채우기 먼저 제거 후 매칭 행만 칠함"},
            },
            "required": ["sheet", "condition_col", "match_value", "color", "target_col0", "target_col1"],
        },
    },
    {
        "name": "sort_live_range",
        "description": "범위를 특정 열 기준으로 정렬한다. Excel 네이티브 정렬이라 행 전체가 서식·수식까지 보존된 채 재배치된다(값만 다시 쓰는 방식이 아님). key_col은 시트 절대 열 인덱스(0-based). 좌표 0-based.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"},
            "key_col": {"type": "integer", "description": "정렬 기준 열(시트 절대 0-based 인덱스)"},
            "ascending": {"type": "boolean", "default": True},
            "has_header": {"type": "boolean", "default": True, "description": "범위 첫 행이 헤더면 true(정렬에서 제외)"}},
            "required": ["sheet", "row0", "col0", "row1", "col1", "key_col"]}},
    {
        "name": "copy_live_range",
        "description": "범위를 다른 위치(같은/다른 시트)로 복사한다. 값·수식·서식이 함께 복사되고 수식 상대참조는 붙여넣는 위치에 맞게 자동 조정된다. 좌표는 0-based.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "원본 시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"},
            "dest_row": {"type": "integer", "description": "붙여넣을 좌상단 행(0-based)"},
            "dest_col": {"type": "integer", "description": "붙여넣을 좌상단 열(0-based)"},
            "dest_sheet": {"description": "붙여넣을 시트(생략 시 원본과 동일)"}},
            "required": ["sheet", "row0", "col0", "row1", "col1", "dest_row", "dest_col"]}},
    {
        "name": "autofit_live",
        "description": "열 너비/행 높이를 내용에 맞춘다(열을 넓혀 한 줄로 다 보이게). 범위를 안 주면 사용 범위 전체. 줄바꿈으로 다 보이게 하려면 format_live_range(action='wrap').",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"}},
            "required": ["sheet"]}},
    {
        "name": "freeze_live_panes",
        "description": "상단 rows개 행 / 좌측 cols개 열을 고정한다(스크롤해도 고정). 헤더 1줄 고정은 rows=1, cols=0. rows=0 이고 cols=0 이면 고정 해제.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "rows": {"type": "integer", "default": 1, "description": "고정할 상단 행 수"},
            "cols": {"type": "integer", "default": 0, "description": "고정할 좌측 열 수"}},
            "required": ["sheet"]}},
    {
        "name": "set_live_autofilter",
        "description": "범위에 자동필터를 켜거나(on=true) 끈다(on=false). ★Windows 전용 - macOS Excel은 미지원이니 정렬(sort_live_range)이나 조건부 강조(color_live_rows_where)로 대체. 좌표 0-based.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"},
            "on": {"type": "boolean", "default": True}},
            "required": ["sheet", "row0", "col0", "row1", "col1"]}},
    {
        "name": "format_live_range",
        "description": (
            "범위에 서식을 적용하는 통합 도구. action으로 종류를 고른다. 좌표는 모두 0-based.\n"
            "- fill_color: 배경색. color 필요('#FFFF00'/'yellow'/'노란색', 'none'이면 제거). "
            "★'어떤 열=값인 행'을 칠하는 조건부 색칠은 이걸 행마다 부르지 말고 color_live_rows_where를 써라.\n"
            "- font: 글꼴. bold(굵게)/font_color(글자색)/color(배경) 중 준 것만 적용.\n"
            "- number_format: 표시 형식. format_code 필요('#,##0'/'0.00%'/'yyyy-mm-dd'/'₩#,##0'/'@'). 값은 안 바뀜.\n"
            "- borders: 테두리. which('all'/'outline'/'top'/'bottom'/'left'/'right')·style('thin'/'medium'/'thick').\n"
            "- wrap: 자동 줄바꿈(긴 내용 다 보이게). autofit_rows=true(기본)면 행 높이도 맞춤.\n"
            "- merge / unmerge: 셀 병합 / 해제.\n"
            "- data_validation: 드롭다운(목록) 유효성. values 필요.\n"
            "- clear: 범위 지우기. contents(값·수식, 기본 true)/formats(서식, 기본 false)."
        ),
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"},
            "action": {"type": "string", "enum": ["fill_color", "font", "number_format", "borders", "wrap", "merge", "unmerge", "data_validation", "clear"]},
            "color": {"description": "fill_color/font의 색('#FFFF00'/'yellow'/'none')"},
            "bold": {"type": "boolean"},
            "font_color": {"description": "font 글자색('#FF0000'/'red' 등)"},
            "format_code": {"type": "string", "description": "number_format 코드"},
            "which": {"type": "string", "enum": ["all", "outline", "top", "bottom", "left", "right"], "default": "all"},
            "style": {"type": "string", "enum": ["thin", "medium", "thick"], "default": "thin"},
            "wrap": {"type": "boolean", "default": True},
            "autofit_rows": {"type": "boolean", "default": True},
            "values": {"type": "array", "items": {"type": "string"}, "description": "data_validation 드롭다운 항목들"},
            "contents": {"type": "boolean", "default": True, "description": "clear: 값·수식 지우기"},
            "formats": {"type": "boolean", "default": False, "description": "clear: 서식 지우기"}},
            "required": ["sheet", "row0", "col0", "row1", "col1", "action"]}},
    {
        "name": "edit_live_rows_columns",
        "description": "행·열을 삽입/삭제한다(원본 비파괴 - 아래·오른쪽 데이터가 수식·서식째 밀리거나 당겨진다). action=insert_rows/delete_rows/insert_columns/delete_columns. index는 0-based(insert는 이 위치 앞에 삽입, delete는 이 위치부터). count는 개수(기본 1).",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "action": {"type": "string", "enum": ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]},
            "index": {"type": "integer", "description": "0-based 행/열 위치"},
            "count": {"type": "integer", "default": 1}},
            "required": ["sheet", "action", "index"]}},
    {
        "name": "set_live_visibility",
        "description": "행 또는 열을 숨기거나 다시 보이게 한다. action=rows/columns, start..end는 0-based(포함), visible=false면 숨김.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "action": {"type": "string", "enum": ["rows", "columns"]},
            "start": {"type": "integer", "description": "시작 행/열(0-based)"},
            "end": {"type": "integer", "description": "끝 행/열(0-based, 포함)"},
            "visible": {"type": "boolean"}},
            "required": ["sheet", "action", "start", "end", "visible"]}},
    {
        "name": "resize_live",
        "description": "열 너비 또는 행 높이를 지정한다. action=column_width(문자 단위)/row_height(포인트 단위). index는 0-based 열/행, size는 값.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "action": {"type": "string", "enum": ["column_width", "row_height"]},
            "index": {"type": "integer", "description": "0-based 열 또는 행 인덱스"},
            "size": {"type": "number"}},
            "required": ["sheet", "action", "index", "size"]}},
    {
        "name": "manage_live_sheet",
        "description": (
            "시트를 관리한다. action으로 고른다.\n"
            "- list: 시트 목록(순서·이름) 반환.\n"
            "- add: 새 시트 추가. name(이름)·before/after(위치, 인덱스나 이름) 선택. 만든 이름 반환.\n"
            "- rename: sheet를 new_name으로.\n"
            "- delete: sheet 삭제(되돌릴 수 없음). ★사용자가 시트 삭제를 명시적으로 요청할 때만.\n"
            "- move: sheet를 before 또는 after 위치로 이동(서식·수식 보존)."
        ),
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"},
            "action": {"type": "string", "enum": ["list", "add", "rename", "delete", "move"]},
            "sheet": {"description": "대상 시트(인덱스 또는 이름). rename/delete/move에 필요"},
            "name": {"type": "string", "description": "add: 새 시트 이름"},
            "new_name": {"type": "string", "description": "rename: 바꿀 이름"},
            "before": {"description": "add/move: 이 시트(인덱스나 이름) 앞"},
            "after": {"description": "add/move: 이 시트(인덱스나 이름) 뒤"}},
            "required": ["action"]}},
    {
        "name": "manage_live_name",
        "description": "이름 정의를 관리한다. action=list(목록)/define(추가, name+refers_to='=Sheet1!$A$1:$A$10' 형식)/delete(삭제, name). 수식에서 이름으로 참조 가능.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"},
            "action": {"type": "string", "enum": ["list", "define", "delete"]},
            "name": {"type": "string"},
            "refers_to": {"type": "string", "description": "define: '=Sheet1!$A$1:$A$10' 형식"}},
            "required": ["action"]}},
    {
        "name": "inspect_csv",
        "description": (
            "CSV 파일의 구조를 조사한다: 인코딩/구분자를 자동 감지하고 크기·헤더· "
            "미리보기 표를 반환한다. CSV를 다루기 전에 항상 먼저 호출해야 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "CSV 파일 절대경로"},
                "preview_rows": {"type": "integer", "default": 12},
                "max_cell_len": {"type": "integer", "default": 60},
                "has_header": {
                    "type": "boolean", "default": True,
                    "description": "첫 행을 헤더로 볼지. 자동판별은 신뢰할 수 없어 기본 true. 헤더가 없는 파일이면 false로 명시.",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "get_csv_cell",
        "description": "CSV 셀 하나의 전체 텍스트를 절단 없이 반환한다(헤더 행 포함 0-based row).",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "row": {"type": "integer"}, "col": {"type": "integer"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["path", "row", "col"],
        },
    },
    {
        "name": "set_csv_cell",
        "description": (
            "CSV 셀 하나를 고치고 파일 전체를 원래 인코딩/구분자로 다시 저장한다. "
            "행/열이 기존 범위 밖이면 빈 칸으로 확장한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "row": {"type": "integer"}, "col": {"type": "integer"},
                "value": {"type": "string"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["path", "row", "col", "value"],
        },
    },
    {
        "name": "append_csv_row",
        "description": "CSV 맨 끝에 행을 추가한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "values": {"type": "array", "items": {"type": "string"}},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["path", "values"],
        },
    },
    {
        "name": "write_csv_table",
        "description": "CSV를 새로 만들거나 통째로 덮어쓴다(데이터 불러오기/내보내기용).",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "header": {"type": "array", "items": {"type": "string"}},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                },
                "encoding": {"type": "string", "default": "utf-8-sig"},
            },
            "required": ["path", "rows"],
        },
    },
    {
        "name": "open_csv_in_excel",
        "description": (
            "CSV를 로컬 Excel에서 실제로 연다. 연 뒤에는 이 파일을 workbook_id로 "
            "얻어 get_live_schema/set_live_cell 등 라이브 도구로 다룰 수 있다. "
            "CSV에 수식을 넣거나 서식을 입히려면 이 경로를 거쳐야 한다(CSV 포맷 "
            "자체엔 수식이 없다)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "open_libreoffice_document",
        "description": (
            "Microsoft Excel이 설치돼 있지 않은 환경(폐쇄망 등)에서 xlsx/ods 파일을 "
            "LibreOffice로 직접 연다. 반환된 path를 이후 모든 libreoffice 도구의 doc "
            "인자로 쓴다. 이미 열려 있는 파일이면 새로 열지 않고 그대로 재사용한다 "
            "(already_open=true). Excel이 설치돼 있다면 대신 open_csv_in_excel/"
            "get_live_schema 계열을 쓰는 게 사용자에게 더 익숙하다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "xlsx 파일 절대경로"},
                "hidden": {
                    "type": "boolean", "default": True,
                    "description": "true면 LibreOffice 창을 화면에 띄우지 않는다. 사용자에게 진행을 보여주려면 false.",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "close_libreoffice_document",
        "description": "open_libreoffice_document로 연 문서를 닫는다. save=true면 닫기 전에 저장한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string", "description": "open_libreoffice_document가 준 path"},
                "save": {"type": "boolean", "default": False},
            },
        },
    },
    {
        "name": "list_open_libreoffice_documents",
        "description": "지금 LibreOffice에 열려 있는 문서 목록. 각 항목의 path가 doc 인자다.",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_libreoffice_schema",
        "description": (
            "LibreOffice로 연 통합문서의 시트 하나를 조사한다: 크기, 미리보기 표, "
            "병합 셀, 수식 존재 여부. 셀을 편집하기 전에 먼저 호출해 레이아웃을 "
            "파악해야 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string", "description": "open_libreoffice_document가 준 path. 생략하면 열려 있는 문서 중 하나가 임의로 선택되니 여러 개 열려 있으면 명시적으로 넘겨야 한다."},
                "sheet": {"description": "시트 인덱스(0-based) 또는 이름", "default": 0},
                "preview_rows": {"type": "integer", "default": 12},
                "max_cell_len": {"type": "integer", "default": 60},
            },
        },
    },
    {
        "name": "get_libreoffice_cell",
        "description": "셀 하나의 값·수식·병합 정보를 절단 없이 전체로 반환한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row": {"type": "integer", "description": "0-based 행"},
                "col": {"type": "integer", "description": "0-based 열"},
            },
            "required": ["sheet", "row", "col"],
        },
    },
    {
        "name": "set_libreoffice_cell",
        "description": (
            "셀 하나를 즉시 갱신하고 파일에 반영한다(디스크 저장은 별도로 "
            "save_libreoffice_document 호출 필요). 숫자처럼 보이는 문자열은 자동으로 "
            "숫자로 기록해 SUM 등 기존 수식이 계속 동작하게 한다. as_formula=true면 "
            "value를 '='로 시작하는 수식으로 그대로 쓴다. 병합 셀의 non-anchor "
            "좌표에 쓰면 오류가 난다. allow_merge_redirect=true로 anchor에 리다이렉트 "
            "하거나 anchor 좌표로 직접 호출한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row": {"type": "integer"},
                "col": {"type": "integer"},
                "value": {"type": "string"},
                "as_formula": {"type": "boolean", "default": False},
                "allow_merge_redirect": {"type": "boolean", "default": False},
            },
            "required": ["sheet", "row", "col", "value"],
        },
    },
    {
        "name": "read_libreoffice_range",
        "description": (
            "직사각형 범위를 값/수식 2차원 배열로 한 번에 읽는다. 셀을 하나씩 "
            "read하는 것보다 훨씬 빠르다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "row1": {"type": "integer"}, "col1": {"type": "integer"},
            },
            "required": ["sheet", "row0", "col0", "row1", "col1"],
        },
    },
    {
        "name": "write_libreoffice_range",
        "description": "직사각형 범위를 한 번에 덮어쓴다(붙여넣기와 동일한 동작). 각 셀 값 앞에 '='를 붙이면 수식으로 해석된다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                    "description": "행 우선 2차원 배열",
                },
            },
            "required": ["sheet", "row0", "col0", "rows"],
        },
    },
    {
        "name": "append_libreoffice_row",
        "description": "지금 사용 범위 바로 아래에 새 행을 추가한다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "values": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["sheet", "values"],
        },
    },
    {
        "name": "save_libreoffice_document",
        "description": "지금까지의 변경을 원본 파일 경로에 그대로 저장한다(형식 유지).",
        "input_schema": {
            "type": "object",
            "properties": {"doc": {"type": "string"}},
        },
    },
]

TOOL_HANDLERS: dict[str, Callable[..., dict[str, Any]]] = {
    "list_open_workbooks": list_open_workbooks,
    "get_workbook_overview": get_workbook_overview,
    "get_live_schema": get_live_schema,
    "get_live_cell": get_live_cell,
    "get_live_table_region": get_live_table_region,
    "set_live_cell": set_live_cell,
    "read_live_range": read_live_range,
    "write_live_range": write_live_range,
    "append_live_row": append_live_row,
    "find_replace_live": find_replace_live,
    "recalculate_live": recalculate_live,
    "activate_live_cell": activate_live_cell,
    "color_live_rows_where": color_live_rows_where,
    "sort_live_range": sort_live_range,
    "copy_live_range": copy_live_range,
    "autofit_live": autofit_live,
    "freeze_live_panes": freeze_live_panes,
    "set_live_autofilter": set_live_autofilter,
    "format_live_range": format_live_range,
    "edit_live_rows_columns": edit_live_rows_columns,
    "set_live_visibility": set_live_visibility,
    "resize_live": resize_live,
    "manage_live_sheet": manage_live_sheet,
    "manage_live_name": manage_live_name,
    "inspect_csv": inspect_csv,
    "get_csv_cell": get_csv_cell,
    "set_csv_cell": set_csv_cell,
    "append_csv_row": append_csv_row,
    "write_csv_table": write_csv_table,
    "open_csv_in_excel": open_csv_in_excel,
    "open_libreoffice_document": open_libreoffice_document,
    "close_libreoffice_document": close_libreoffice_document,
    "list_open_libreoffice_documents": list_open_libreoffice_documents,
    "get_libreoffice_schema": get_libreoffice_schema,
    "get_libreoffice_cell": get_libreoffice_cell,
    "set_libreoffice_cell": set_libreoffice_cell,
    "read_libreoffice_range": read_libreoffice_range,
    "write_libreoffice_range": write_libreoffice_range,
    "append_libreoffice_row": append_libreoffice_row,
    "save_libreoffice_document": save_libreoffice_document,
}


def call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """이름으로 tool 실행. 예외도 dict로 직렬화한다."""
    handler = TOOL_HANDLERS.get(name)
    if handler is None:
        return {"error": f"unknown tool: {name}"}
    try:
        return handler(**arguments)
    except SeepageError as e:
        return {"error": type(e).__name__, "message": str(e)}
    except (IndexError, ValueError, FileNotFoundError) as e:
        return {"error": type(e).__name__, "message": str(e)}
    except Exception as e:
        return {"error": "unexpected", "type": type(e).__name__, "message": str(e)}
