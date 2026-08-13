"""xgen-seepage 공통 타입.

dataclass 명명 규칙은 PlateerLab/document-adapter의 `base.py`를 그대로 따른다 —
같은 XGEN 세션에서 두 MCP 서버(document-adapter: 닫힌 파일 편집,
xgen-seepage: 열려 있는 통합문서 실시간 편집)의 도구를 함께 쓰는 LLM이
스키마마다 다른 관례를 재학습하지 않도록 하기 위함이다. 근거는 저장소
루트 `ARCHITECTURE.md` 및 `NOTICE` 참조.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


class SeepageError(Exception):
    """xgen-seepage 공통 베이스 예외."""


class WorkbookNotFoundError(SeepageError):
    """지정한 workbook_id로 열려 있는 통합문서를 찾을 수 없음(창이 닫혔을 수 있음)."""


class SheetIndexError(SeepageError, IndexError):
    """시트 인덱스/이름이 통합문서에 없음."""


class CellOutOfBoundsError(SeepageError, IndexError):
    """셀 좌표가 사용 범위를 벗어남."""


class MergedCellWriteError(SeepageError, ValueError):
    """병합 셀의 non-anchor 좌표에 직접 쓰기를 시도함. anchor로 쓰거나
    allow_merge_redirect=True를 전달해야 한다."""


class ExcelUnavailableError(SeepageError):
    """xlwings 미설치 또는 로컬에 연결 가능한 Excel 애플리케이션이 없음."""


class RangeTooLargeError(SeepageError, ValueError):
    """요청한 범위가 안전 상한을 넘음 — COM 호출 폭주/컨텍스트 폭발 방지."""


@dataclass
class MergeInfo:
    anchor: tuple[int, int]
    span: tuple[int, int]

    def to_dict(self) -> dict[str, Any]:
        return {"anchor": list(self.anchor), "span": list(self.span)}


@dataclass
class CellContent:
    """셀 하나의 전체 정보. `value`는 계산된 값(수식이면 그 결과),
    `formula`는 수식 문자열(수식이 아니면 None) — 두 필드를 분리해
    "무엇이 적혀 있고 무엇으로 계산됐는지"를 LLM이 혼동하지 않게 한다."""

    row: int
    col: int
    value: Any
    formula: str | None
    text: str
    is_anchor: bool
    anchor: tuple[int, int]
    span: tuple[int, int]

    def to_dict(self) -> dict[str, Any]:
        return {
            "row": self.row,
            "col": self.col,
            "value": self.value,
            "formula": self.formula,
            "text": self.text,
            "is_anchor": self.is_anchor,
            "anchor": list(self.anchor),
            "span": list(self.span),
        }


@dataclass
class SheetSchema:
    index: int
    name: str
    rows: int
    cols: int
    preview: list[list[str | None]]
    merges: list[MergeInfo] = field(default_factory=list)
    has_formulas: bool = False
    truncated: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "name": self.name,
            "rows": self.rows,
            "cols": self.cols,
            "preview": self.preview,
            "merges": [m.to_dict() for m in self.merges],
            "has_formulas": self.has_formulas,
            "truncated": self.truncated,
        }


@dataclass
class WorkbookInfo:
    workbook_id: str
    """`<app_pid>:<book_name>` 합성 id. 같은 이름의 통합문서가 여러 Excel
    인스턴스에 열려 있어도 안정적으로 구분하기 위함."""
    name: str
    full_path: str | None
    app_pid: int
    sheets: list[str]
    active_sheet: str | None
    saved: bool

    def to_dict(self) -> dict[str, Any]:
        return {
            "workbook_id": self.workbook_id,
            "name": self.name,
            "full_path": self.full_path,
            "app_pid": self.app_pid,
            "sheets": self.sheets,
            "active_sheet": self.active_sheet,
            "saved": self.saved,
        }
