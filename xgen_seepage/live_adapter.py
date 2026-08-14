"""라이브 어댑터: xlwings 기반, "지금 열려 있는" 통합문서를 그 자리에서 편집한다.

이것이 xgen-seepage의 핵심 차별점이다. document-adapter(PlateerLab)는 **디스크의
파일**을 openpyxl로 열고/고치고/저장하는 배치형 도구다. Claude for Excel /
ChatGPT for Excel처럼 "지금 화면에 떠 있는 시트를 실시간으로 편집"하려면
파일 저장을 거치지 않고 **살아 있는 Excel 프로세스**에 직접 붙어야 한다.

두 가지 접근을 비교해 xlwings(로컬 COM/AppleScript 자동화)를 선택했다:
  - Office.js 태스크팬 방식(Claude/ChatGPT for Excel 실제 구현): HTTPS로 호스팅되는
    웹앱이 필요하다. `localhost` 인증서로 오프라인 자체는 가능하지만, 매니페스트
    사이드로딩·인증서 신뢰·SharedRuntime 설정 등 설치 마찰이 크다.
  - xlwings COM 자동화(이 모듈의 접근): `pip install`만으로 끝난다. 웹서버도
    인증서도 매니페스트도 없다. 순수 로컬 프로세스 간 통신이라 인터넷 요구사항이
    원천적으로 없다. 대신 Windows/macOS 전용이고 로컬에 Excel 설치가 필요하다.
자세한 리서치 근거는 저장소 루트 `ARCHITECTURE.md` 참조.

병합 셀 판정과 숫자/텍스트 변환 휴리스틱은 PlateerLab/document-adapter의
`xlsx_adapter.py`와 동일 시맨틱을 따른다(NOTICE 참조). 대상이 파일이 아니라
xlwings Range라는 점만 다르다.

⚠️ 성능: 각 셀 접근이 COM 왕복 1회다. 시트 전체를 셀 단위로 훑는 것은 느리다.
그래서 값 읽기는 `Range.value`(2차원 리스트 벌크 반환)를 쓰고, 병합 셀 유무는
`UsedRange.MergeCells`로 먼저 빠르게 판정한 뒤(병합이 아예 없으면 O(1)),
있을 때만 셀 단위로 위치를 확인한다. 그마저도 `_MAX_MERGE_SCAN_CELLS`로
상한을 두고 넘으면 `truncated=True`로 알린다. 크면 range를 좁혀 다시 부르라는
신호다(자세한 근거는 이 파일의 `_merge_map` 참고).
"""
from __future__ import annotations

from typing import Any

from ._cellfmt import cell_text, coerce_for_write
from .base import (
    CellContent,
    CellOutOfBoundsError,
    ExcelUnavailableError,
    MergeInfo,
    MergedCellWriteError,
    RangeTooLargeError,
    SheetIndexError,
    SheetSchema,
    WorkbookInfo,
    WorkbookNotFoundError,
)

try:
    import xlwings as xw
except ImportError:  # xlwings는 Windows/macOS 전용. 미설치 환경에서도 import 자체는 성공해야 함
    xw = None  # type: ignore[assignment]

_MAX_MERGE_SCAN_CELLS = 20_000
_MAX_RANGE_CELLS = 50_000


def _require_xlwings() -> Any:
    if xw is None:
        raise ExcelUnavailableError(
            "xlwings가 설치되어 있지 않습니다. `pip install xgen-seepage[live]`로 "
            "설치하세요. 라이브 모드는 Windows 또는 macOS에서 로컬 Excel 실행이 "
            "필요합니다(오프라인 동작. 인터넷 불필요)."
        )
    return xw


def _workbook_id(app: Any, book: Any) -> str:
    return f"{app.pid}:{book.name}"


def list_open_workbooks() -> list[WorkbookInfo]:
    """로컬에서 실행 중인 모든 Excel 인스턴스의 열린 통합문서 목록.

    LLM이 어떤 파일을 대상으로 할지 고를 때 항상 먼저 호출해야 한다.
    workbook_id를 생략하면 이후 호출은 "현재 활성 통합문서"를 임의로 집는다.
    """
    xw_ = _require_xlwings()
    out: list[WorkbookInfo] = []
    for app in xw_.apps:
        try:
            active = app.books.active
            active_name = active.name if active else None
        except Exception:
            active_name = None
        for book in app.books:
            out.append(
                WorkbookInfo(
                    workbook_id=_workbook_id(app, book),
                    name=book.name,
                    full_path=book.fullname if book.fullname != book.name else None,
                    app_pid=app.pid,
                    sheets=[s.name for s in book.sheets],
                    active_sheet=active_name,
                    saved=book.saved,
                )
            )
    return out


def _resolve_book(workbook_id: str | None) -> Any:
    xw_ = _require_xlwings()
    if workbook_id is None:
        for app in xw_.apps:
            book = app.books.active
            if book is not None:
                return book
        raise WorkbookNotFoundError(
            "열려 있는 통합문서가 없습니다. Excel에서 파일을 먼저 열어주세요."
        )
    pid_s, sep, name = workbook_id.partition(":")
    if not sep:
        raise WorkbookNotFoundError(f"잘못된 workbook_id 형식: {workbook_id!r}")
    try:
        pid = int(pid_s)
    except ValueError:
        raise WorkbookNotFoundError(f"잘못된 workbook_id: {workbook_id!r}") from None
    for app in xw_.apps:
        if app.pid == pid:
            for book in app.books:
                if book.name == name:
                    return book
    raise WorkbookNotFoundError(
        f"workbook_id {workbook_id!r}를 찾을 수 없습니다(창이 닫혔을 수 있음). "
        "list_open_workbooks로 다시 조회하세요."
    )


def _sheet(workbook_id: str | None, sheet: int | str) -> Any:
    book = _resolve_book(workbook_id)
    try:
        return book.sheets[sheet]
    except Exception as exc:
        raise SheetIndexError(f"sheet {sheet!r} not found in {book.name!r}") from exc


def _used_dims(ws: Any) -> tuple[int, int]:
    ur = ws.used_range
    if ur is None:
        return 0, 0
    last = ur.last_cell
    return last.row, last.column  # 1-based counts == rows/cols


def _merge_map(ws: Any, rows: int, cols: int) -> tuple[dict, dict, bool]:
    """(anchor→span), (covered→anchor), truncated. 좌표는 0-based.

    1단계: `UsedRange.MergeCells`로 병합 유무를 한 번의 COM 호출로 판정한다
    (False = 병합 전혀 없음 → 즉시 빈 결과 반환, 압도적으로 흔한 경우).
    2단계: 병합이 있으면(True 또는 tri-state None=혼합) 셀 단위로 스캔하되
    `_MAX_MERGE_SCAN_CELLS`를 넘으면 스캔을 멈추고 truncated=True를 반환한다.
    무한정 COM 왕복을 하는 대신, 큰 시트는 범위를 좁혀 다시 조회하라는
    신호를 준다.
    """
    if rows <= 0 or cols <= 0:
        return {}, {}, False
    used = ws.range((1, 1), (rows, cols))
    try:
        has_merge = used.api.MergeCells
    except Exception:
        has_merge = None
    if has_merge is False:
        return {}, {}, False

    anchors: dict[tuple[int, int], tuple[int, int]] = {}
    covered: dict[tuple[int, int], tuple[int, int]] = {}
    seen: set[tuple[int, int]] = set()
    truncated = False
    scanned = 0
    for r in range(rows):
        for c in range(cols):
            if (r, c) in seen:
                continue
            scanned += 1
            if scanned > _MAX_MERGE_SCAN_CELLS:
                truncated = True
                return anchors, covered, truncated
            cell_api = ws.range((r + 1, c + 1)).api
            area = getattr(cell_api, "MergeArea", None)
            if area is None:
                continue
            if area.Rows.Count <= 1 and area.Columns.Count <= 1:
                continue
            ar0, ac0 = area.Row - 1, area.Column - 1
            span = (area.Rows.Count, area.Columns.Count)
            anchors[(ar0, ac0)] = span
            for rr in range(ar0, ar0 + span[0]):
                for cc in range(ac0, ac0 + span[1]):
                    seen.add((rr, cc))
                    if (rr, cc) != (ar0, ac0):
                        covered[(rr, cc)] = (ar0, ac0)
    return anchors, covered, truncated


def get_sheet_schema(
    workbook_id: str | None,
    sheet: int | str,
    *,
    preview_rows: int = 12,
    max_cell_len: int = 60,
) -> SheetSchema:
    """시트 하나의 크기·미리보기·병합 정보. 항상 첫 호출로 사용."""
    ws = _sheet(workbook_id, sheet)
    rows, cols = _used_dims(ws)
    visible = min(rows, preview_rows)
    anchors, covered, truncated = _merge_map(ws, rows, cols)

    preview: list[list[str | None]] = [[None] * cols for _ in range(visible)]
    if visible and cols:
        block = ws.range((1, 1), (visible, cols)).options(ndim=2).value
        formula_block = ws.range((1, 1), (visible, cols)).formula
        for r in range(visible):
            for c in range(cols):
                if (r, c) in covered:
                    continue
                raw = block[r][c] if visible > 0 else None
                formula = formula_block[r][c] if isinstance(formula_block, list) else formula_block
                text = cell_text(raw)
                if isinstance(formula, str) and formula.startswith("="):
                    text = f"{cell_text(raw)}"  # 계산된 값을 보여준다(수식은 get_cell에서 확인)
                preview[r][c] = text[:max_cell_len]

    has_formulas = False
    if rows and cols:
        try:
            f_block = ws.range((1, 1), (rows, cols)).formula
            if isinstance(f_block, str):
                has_formulas = f_block.startswith("=")
            else:
                has_formulas = any(
                    isinstance(v, str) and v.startswith("=")
                    for row_ in f_block
                    for v in row_
                )
        except Exception:
            has_formulas = False

    return SheetSchema(
        index=ws.index - 1,
        name=ws.name,
        rows=rows,
        cols=cols,
        preview=preview,
        merges=[MergeInfo(anchor=a, span=s) for a, s in anchors.items()],
        has_formulas=has_formulas,
        truncated=truncated,
    )


def get_cell(workbook_id: str | None, sheet: int | str, row: int, col: int) -> CellContent:
    """셀 하나의 값·수식·병합 정보를 전부 반환한다(미리보기의 40~60자 절단 없이)."""
    ws = _sheet(workbook_id, sheet)
    rows, cols = _used_dims(ws)
    if row < 0 or col < 0 or row >= max(rows, row + 1) or col >= max(cols, col + 1):
        # 사용 범위 밖도 read는 허용한다(빈 셀 조회는 자연스러운 요청). 다만 완전히
        # 음수거나 시트 범위(1,048,576행 / 16,384열)를 넘으면 명백한 오류로 거절.
        if row < 0 or col < 0 or row >= 1_048_576 or col >= 16_384:
            raise CellOutOfBoundsError(f"cell ({row},{col}) is out of the sheet's valid range")

    anchors, covered, _ = _merge_map(ws, max(rows, row + 1), max(cols, col + 1))
    if (row, col) in covered:
        anchor = covered[(row, col)]
        span = anchors[anchor]
        is_anchor = False
    else:
        anchor = (row, col)
        span = anchors.get((row, col), (1, 1))
        is_anchor = True

    rng = ws.range((anchor[0] + 1, anchor[1] + 1))
    value = rng.value
    formula = rng.formula if isinstance(rng.formula, str) and rng.formula.startswith("=") else None
    text = cell_text(value)
    return CellContent(
        row=row, col=col, value=value, formula=formula, text=text,
        is_anchor=is_anchor, anchor=anchor, span=span,
    )


def _resolve_writable(
    ws: Any, row: int, col: int, allow_merge_redirect: bool
) -> tuple[int, int]:
    rows, cols = _used_dims(ws)
    _, covered, _ = _merge_map(ws, max(rows, row + 1), max(cols, col + 1))
    if (row, col) in covered:
        if not allow_merge_redirect:
            ar, ac = covered[(row, col)]
            raise MergedCellWriteError(
                f"cell ({row},{col}) is part of a merge anchored at ({ar},{ac}). "
                "Write to the anchor, or pass allow_merge_redirect=True."
            )
        return covered[(row, col)]
    return row, col


def set_cell(
    workbook_id: str | None,
    sheet: int | str,
    row: int,
    col: int,
    value: str,
    *,
    as_formula: bool = False,
    allow_merge_redirect: bool = False,
) -> str:
    """살아 있는 시트의 셀 하나를 즉시 갱신한다(파일 저장 불필요. 사용자가
    화면에서 바로 확인한다). `as_formula=True`면 `value`를 수식 문자열
    (`=A1+B1`)로 그대로 쓴다. 아니면 숫자처럼 보이는 문자열은 숫자로,
    나머지는 문자로 자동 판단해서 쓴다(_cellfmt.coerce_for_write)."""
    ws = _sheet(workbook_id, sheet)
    wr, wc = _resolve_writable(ws, row, col, allow_merge_redirect)
    rng = ws.range((wr + 1, wc + 1))
    old = cell_text(rng.value)
    if as_formula:
        if not value.startswith("="):
            raise ValueError("as_formula=True requires a value starting with '='")
        rng.formula = value
    else:
        rng.value = coerce_for_write(value)
    return old


def append_row(workbook_id: str | None, sheet: int | str, values: list[str]) -> int:
    """사용 범위 바로 아래에 새 행을 추가한다. 반환값은 0-based 행 인덱스."""
    ws = _sheet(workbook_id, sheet)
    rows, _ = _used_dims(ws)
    at_row = rows  # 0-based 다음 행
    for c, v in enumerate(values):
        if v == "":
            continue
        ws.range((at_row + 1, c + 1)).value = coerce_for_write(str(v))
    return at_row


def read_range(
    workbook_id: str | None,
    sheet: int | str,
    row0: int,
    col0: int,
    row1: int,
    col1: int,
) -> dict[str, Any]:
    """직사각형 범위를 값/수식 2차원 배열로 벌크 반환한다(셀 단위 호출보다
    훨씬 빠름. Range.value/Range.formula는 COM 호출 1회로 전체를 가져온다)."""
    if row1 < row0 or col1 < col0:
        raise ValueError("row1/col1 must be >= row0/col0")
    n_cells = (row1 - row0 + 1) * (col1 - col0 + 1)
    if n_cells > _MAX_RANGE_CELLS:
        raise RangeTooLargeError(
            f"requested range has {n_cells} cells, over the {_MAX_RANGE_CELLS} cap. "
            "Narrow the range and call again."
        )
    ws = _sheet(workbook_id, sheet)
    rng = ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1))
    values = rng.options(ndim=2).value
    formulas = rng.formula
    if not isinstance(formulas, list):
        formulas = [[formulas]]
    return {
        "row0": row0, "col0": col0, "row1": row1, "col1": col1,
        "values": values,
        "formulas": formulas,
    }


def write_range(
    workbook_id: str | None,
    sheet: int | str,
    row0: int,
    col0: int,
    rows: list[list[str]],
) -> None:
    """직사각형 범위를 한 번에 덮어쓴다. 각 셀은 set_cell과 동일한 숫자/문자
    휴리스틱을 적용한다(수식을 쓰려면 셀 값 앞에 '='를 붙이면 된다. Excel
    자체가 '='로 시작하는 문자열 대입을 수식으로 해석한다)."""
    if not rows:
        return
    n_cols = max(len(r) for r in rows)
    n_cells = len(rows) * n_cols
    if n_cells > _MAX_RANGE_CELLS:
        raise RangeTooLargeError(
            f"write covers {n_cells} cells, over the {_MAX_RANGE_CELLS} cap. "
            "Split into smaller writes."
        )
    ws = _sheet(workbook_id, sheet)
    coerced: list[list[Any]] = []
    for r in rows:
        line: list[Any] = []
        for v in r:
            s = str(v)
            line.append(s if s.startswith("=") else coerce_for_write(s))
        coerced.append(line)
    rng = ws.range((row0 + 1, col0 + 1), (row0 + len(rows), col0 + n_cols))
    rng.value = coerced


def activate(workbook_id: str | None, sheet: int | str | None = None,
             row: int | None = None, col: int | None = None) -> None:
    """Excel 창을 앞으로 가져오고, 필요하면 시트를 전환하고 셀을 선택한다.

    기능적으로는 필요 없지만, 에이전트가 "지금 이 셀을 보고 있다"는 걸
    사용자가 눈으로 따라갈 수 있게 한다. Claude/ChatGPT for Excel과 같은
    체감을 위한 UX 보조 도구다.
    """
    book = _resolve_book(workbook_id)
    book.app.activate(steal_focus=True)
    ws = book.sheets[sheet] if sheet is not None else book.sheets.active
    ws.activate()
    if row is not None and col is not None:
        ws.range((row + 1, col + 1)).select()
