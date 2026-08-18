"""MCP / Anthropic tool-use 정의 + dispatcher.

스키마 형태(각 tool = {name, description, input_schema} dict, dispatcher는
이름→dict 반환 함수, 예외는 {"error":..., "message":...}로 직렬화)는
PlateerLab/document-adapter의 `tools.py` 관례를 그대로 따른다. 같은 XGEN
세션에서 document-adapter(닫힌 파일)와 xgen-seepage(열린 통합문서) 도구를
같이 등록해도 LLM이 두 서버를 다른 스타일로 다룰 필요가 없게 하기 위함.
"""
from __future__ import annotations

from typing import Any, Callable

from . import csv_adapter, libreoffice_adapter, live_adapter
from .base import SeepageError

# -------- 라이브(열려 있는 통합문서) 도구 --------


def list_open_workbooks() -> dict[str, Any]:
    books = live_adapter.list_open_workbooks()
    return {"workbooks": [b.to_dict() for b in books]}


def get_live_schema(workbook_id: str | None = None, sheet: int | str = 0,
                     preview_rows: int = 12, max_cell_len: int = 60) -> dict[str, Any]:
    schema = live_adapter.get_sheet_schema(
        workbook_id, sheet, preview_rows=preview_rows, max_cell_len=max_cell_len)
    return schema.to_dict()


def get_live_cell(workbook_id: str | None, sheet: int | str, row: int, col: int) -> dict[str, Any]:
    return live_adapter.get_cell(workbook_id, sheet, row, col).to_dict()


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


def activate_live_cell(workbook_id: str | None = None, sheet: int | str | None = None,
                        row: int | None = None, col: int | None = None) -> dict[str, Any]:
    live_adapter.activate(workbook_id, sheet, row, col)
    return {"ok": True}


def list_live_sheets(workbook_id: str | None = None) -> dict[str, Any]:
    return {"sheets": live_adapter.list_sheets(workbook_id)}


def add_live_sheet(workbook_id: str | None = None, name: str | None = None,
                    before: int | str | None = None, after: int | str | None = None) -> dict[str, Any]:
    return {"name": live_adapter.add_sheet(workbook_id, name, before, after)}


def rename_live_sheet(workbook_id: str | None, sheet: int | str, new_name: str) -> dict[str, Any]:
    return {"name": live_adapter.rename_sheet(workbook_id, sheet, new_name)}


def delete_live_sheet(workbook_id: str | None, sheet: int | str) -> dict[str, Any]:
    live_adapter.delete_sheet(workbook_id, sheet)
    return {"ok": True}


def move_live_sheet(workbook_id: str | None, sheet: int | str,
                     before: int | str | None = None, after: int | str | None = None) -> dict[str, Any]:
    live_adapter.move_sheet(workbook_id, sheet, before, after)
    return {"ok": True}


def set_live_fill_color(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                        color: Any, workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.set_fill_color(workbook_id, sheet, row0, col0, row1, col1, color)
    return {"ok": True}


def format_live_range(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                      fill_color: Any = None, bold: bool | None = None, font_color: Any = None,
                      workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.format_range(workbook_id, sheet, row0, col0, row1, col1, fill_color, bold, font_color)
    return {"ok": True}


def set_live_number_format(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                           format_code: str, workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.set_number_format(workbook_id, sheet, row0, col0, row1, col1, format_code)
    return {"ok": True}


def merge_live_cells(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                     workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.merge_cells(workbook_id, sheet, row0, col0, row1, col1)
    return {"ok": True}


def unmerge_live_cells(sheet: int | str, row0: int, col0: int, row1: int, col1: int,
                       workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.unmerge_cells(workbook_id, sheet, row0, col0, row1, col1)
    return {"ok": True}


def autofit_live(sheet: int | str, row0: int | None = None, col0: int | None = None,
                 row1: int | None = None, col1: int | None = None,
                 workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.autofit(workbook_id, sheet, row0, col0, row1, col1)
    return {"ok": True}


def set_live_column_width(sheet: int | str, col: int, width: float,
                          workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.set_column_width(workbook_id, sheet, col, width)
    return {"ok": True}


def set_live_row_height(sheet: int | str, row: int, height: float,
                        workbook_id: str | None = None) -> dict[str, Any]:
    live_adapter.set_row_height(workbook_id, sheet, row, height)
    return {"ok": True}


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
            "생략하면 이후 호출은 '현재 활성 통합문서'를 임의로 집는다."
        ),
        "input_schema": {"type": "object", "properties": {}},
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
        "name": "write_live_range",
        "description": (
            "직사각형 범위를 한 번에 덮어쓴다(붙여넣기와 동일한 동작). "
            "각 셀 값 앞에 '='를 붙이면 수식으로 해석된다. 최대 50,000셀."
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
        "name": "list_live_sheets",
        "description": "열려 있는 통합문서의 시트 목록(순서·이름, 0-based index).",
        "input_schema": {
            "type": "object",
            "properties": {"workbook_id": {"type": "string"}},
        },
    },
    {
        "name": "add_live_sheet",
        "description": (
            "통합문서에 새 시트를 추가한다. name으로 이름 지정 가능, "
            "before/after(시트 인덱스나 이름)로 위치 지정. 안 주면 맨 뒤에 추가한다. "
            "만들어진 시트 이름을 반환."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "name": {"type": "string", "description": "새 시트 이름(생략 시 기본 이름)"},
                "before": {"description": "이 시트(인덱스 또는 이름) 앞에 삽입"},
                "after": {"description": "이 시트(인덱스 또는 이름) 뒤에 삽입"},
            },
        },
    },
    {
        "name": "rename_live_sheet",
        "description": "시트 이름을 바꾼다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "new_name": {"type": "string"},
            },
            "required": ["sheet", "new_name"],
        },
    },
    {
        "name": "delete_live_sheet",
        "description": "시트를 삭제한다(마지막 한 장은 Excel이 삭제를 막는다).",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
            },
            "required": ["sheet"],
        },
    },
    {
        "name": "move_live_sheet",
        "description": "시트를 다른 위치로 이동(재정렬)한다. before 또는 after로 대상 시트(인덱스나 이름)를 지정.",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "이동할 시트(인덱스 또는 이름)"},
                "before": {"description": "이 시트 앞으로 이동"},
                "after": {"description": "이 시트 뒤로 이동"},
            },
            "required": ["sheet"],
        },
    },
    {
        "name": "set_live_fill_color",
        "description": (
            "셀 범위의 배경(채우기) 색을 칠한다. color는 '#FFFF00' 또는 'yellow'/'노란색' "
            "같은 값. color를 'none'으로 주면 채우기 제거. 조건에 맞는 여러 행을 칠하려면 "
            "각 행 범위(row0=row1=그 행)로 여러 번 호출한다. 예: '결함상태'가 '결함조치완료'인 "
            "행을 노란색으로."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "row1": {"type": "integer"}, "col1": {"type": "integer"},
                "color": {"description": "'#FFFF00' / 'yellow' / '노란색' 등, 또는 'none'(제거)"},
            },
            "required": ["sheet", "row0", "col0", "row1", "col1", "color"],
        },
    },
    {
        "name": "format_live_range",
        "description": "범위에 채우기 색/굵게/글자색을 한 번에 적용한다. 지정한 것만 바꾼다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "workbook_id": {"type": "string"},
                "sheet": {"description": "시트 인덱스 또는 이름"},
                "row0": {"type": "integer"}, "col0": {"type": "integer"},
                "row1": {"type": "integer"}, "col1": {"type": "integer"},
                "fill_color": {"description": "배경색('#FFFF00'/'yellow'/'none')"},
                "bold": {"type": "boolean"},
                "font_color": {"description": "글자색('#FF0000'/'red' 등)"},
            },
            "required": ["sheet", "row0", "col0", "row1", "col1"],
        },
    },
    {
        "name": "set_live_number_format",
        "description": "범위의 표시 형식만 바꾼다(값은 그대로). 예: '#,##0'(천단위), '0.00%'(백분율), 'yyyy-mm-dd'(날짜), '₩#,##0'(원화), '@'(텍스트).",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"},
            "format_code": {"type": "string"}},
            "required": ["sheet", "row0", "col0", "row1", "col1", "format_code"]}},
    {
        "name": "merge_live_cells",
        "description": "범위를 하나로 병합한다(제목/헤더용).",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"}},
            "required": ["sheet", "row0", "col0", "row1", "col1"]}},
    {
        "name": "unmerge_live_cells",
        "description": "병합을 해제한다.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"}},
            "required": ["sheet", "row0", "col0", "row1", "col1"]}},
    {
        "name": "autofit_live",
        "description": "열 너비/행 높이를 내용에 맞춘다. 범위를 안 주면 사용 범위 전체.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row0": {"type": "integer"}, "col0": {"type": "integer"},
            "row1": {"type": "integer"}, "col1": {"type": "integer"}},
            "required": ["sheet"]}},
    {
        "name": "set_live_column_width",
        "description": "열 너비를 지정한다(문자 단위). col은 0-based 열 인덱스.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "col": {"type": "integer"}, "width": {"type": "number"}},
            "required": ["sheet", "col", "width"]}},
    {
        "name": "set_live_row_height",
        "description": "행 높이를 지정한다(포인트 단위). row는 0-based 행 인덱스.",
        "input_schema": {"type": "object", "properties": {
            "workbook_id": {"type": "string"}, "sheet": {"description": "시트 인덱스 또는 이름"},
            "row": {"type": "integer"}, "height": {"type": "number"}},
            "required": ["sheet", "row", "height"]}},
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
    "get_live_schema": get_live_schema,
    "get_live_cell": get_live_cell,
    "set_live_cell": set_live_cell,
    "read_live_range": read_live_range,
    "write_live_range": write_live_range,
    "append_live_row": append_live_row,
    "activate_live_cell": activate_live_cell,
    "list_live_sheets": list_live_sheets,
    "add_live_sheet": add_live_sheet,
    "rename_live_sheet": rename_live_sheet,
    "delete_live_sheet": delete_live_sheet,
    "move_live_sheet": move_live_sheet,
    "set_live_fill_color": set_live_fill_color,
    "format_live_range": format_live_range,
    "set_live_number_format": set_live_number_format,
    "merge_live_cells": merge_live_cells,
    "unmerge_live_cells": unmerge_live_cells,
    "autofit_live": autofit_live,
    "set_live_column_width": set_live_column_width,
    "set_live_row_height": set_live_row_height,
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
    """이름으로 tool 실행. 예외도 dict로 직렬화(document-adapter와 동일 관례)."""
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
