"""live_adapter는 xlwings/실제 Excel 없이도 이 스위트가 도는 CI/개발머신에서
"우아하게 실패"해야 한다. 이 저장소를 빌드한 SERVER_JS에는 Excel이 설치돼
있지 않아, 아래 테스트들이 검증하는 경로가 바로 그 실제 운영 환경이다.
xlwings가 설치돼 있고 정상 import된다는 전제 하에, "실행 중인 Excel이 없을
때" 도구들이 명확한 에러를 내는지 확인한다(조용히 죽거나 원인불명 COM 예외를
그대로 흘리지 않는지)."""
from __future__ import annotations

import pytest

from xgen_seepage import live_adapter, tools
from xgen_seepage.base import ExcelUnavailableError, WorkbookNotFoundError


def test_list_open_workbooks_empty_without_running_excel() -> None:
    # xlwings는 설치돼 있지만 이 머신엔 Excel 자체가 없다. 빈 목록이 되어야
    # 하며, COM 예외가 새어나오면 안 된다.
    books = live_adapter.list_open_workbooks()
    assert books == []


def test_resolve_book_raises_when_nothing_open() -> None:
    with pytest.raises(WorkbookNotFoundError):
        live_adapter._resolve_book(None)


def test_get_sheet_schema_raises_when_nothing_open() -> None:
    with pytest.raises(WorkbookNotFoundError):
        live_adapter.get_sheet_schema(None, 0)


def test_unknown_workbook_id_raises() -> None:
    with pytest.raises(WorkbookNotFoundError):
        live_adapter._resolve_book("999999:no-such-book.xlsx")


def test_malformed_workbook_id_raises() -> None:
    with pytest.raises(WorkbookNotFoundError):
        live_adapter._resolve_book("not-a-valid-id")


def test_excel_unavailable_error_when_xlwings_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(live_adapter, "xw", None)
    with pytest.raises(ExcelUnavailableError):
        live_adapter.list_open_workbooks()


def test_tools_dispatch_wraps_workbook_not_found() -> None:
    result = tools.call_tool("get_live_schema", {"sheet": 0})
    assert result["error"] == "WorkbookNotFoundError"


def test_tool_definitions_match_handlers() -> None:
    names_in_defs = {t["name"] for t in tools.TOOL_DEFINITIONS}
    names_in_handlers = set(tools.TOOL_HANDLERS.keys())
    assert names_in_defs == names_in_handlers
