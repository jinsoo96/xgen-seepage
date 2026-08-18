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

import platform
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


_MAC_TCC_MSG = (
    "macOS 자동화 권한이 없어 Excel을 제어할 수 없습니다. [시스템 설정 > 개인정보 "
    "보호 및 보안 > 자동화]에서, xgen-seepage(또는 이걸 실행한 터미널 앱)가 "
    "'Microsoft Excel' 제어를 허용하도록 켠 뒤 다시 시도하세요. 처음 실행할 때 "
    "권한 요청 창이 뜨면 '확인'을 누르면 됩니다."
)


def _translate_excel_error(exc: Exception) -> None:
    """macOS 자동화 권한 거부(Apple Event 오류 -1743)를 실행 가능한 안내로 바꿔
    다시 던진다. 그 외 예외엔 아무것도 안 한다(호출자가 원래 예외를 그대로
    올린다). 권한이 없으면 xlwings가 Excel에 Apple Event를 보내는 첫 순간부터
    raw appscript CommandError로 막히는데, 알아볼 수 없는 트레이스백 대신 뭘
    해야 하는지 알려주기 위함이다(macOS 최대 함정)."""
    s = str(exc).lower()
    if "-1743" in s or "declined permission" in s or "not permitted" in s:
        raise ExcelUnavailableError(_MAC_TCC_MSG) from exc


def _workbook_id(app: Any, book: Any) -> str:
    return f"{app.pid}:{book.name}"


def _book_saved(book: Any) -> bool:
    """통합문서에 저장되지 않은 변경이 없는지(=Saved). xlwings Book 래퍼엔
    `.saved`가 없어 raw 네이티브 객체(`book.api`)로 접근하는데, 그 속성이
    플랫폼마다 다르다: Windows COM은 `Saved`(bool 속성), macOS는 xlwings가
    appscript 참조를 주므로 `saved`(참조 → `.get()`)다. 어느 쪽도 안 되면
    안전하게 False(변경 있음으로 간주)로 떨어진다 - 이 값은 정보용이라
    core 편집 동작에는 영향이 없다."""
    api = book.api
    for getter in (
        lambda: api.Saved,        # Windows COM
        lambda: api.saved.get(),  # macOS appscript 참조
        lambda: api.saved,        # 혹시 값으로 감싸주는 경우
    ):
        try:
            return bool(getter())
        except Exception:
            continue
    return False


def list_open_workbooks() -> list[WorkbookInfo]:
    """로컬에서 실행 중인 모든 Excel 인스턴스의 열린 통합문서 목록.

    LLM이 어떤 파일을 대상으로 할지 고를 때 항상 먼저 호출해야 한다.
    workbook_id를 생략하면 이후 호출은 "현재 활성 통합문서"를 임의로 집는다.
    """
    xw_ = _require_xlwings()
    out: list[WorkbookInfo] = []
    try:
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
                        saved=_book_saved(book),
                    )
                )
    except Exception as e:
        _translate_excel_error(e)  # macOS 권한 거부면 안내로 바꿔 던진다
        raise
    return out


def _resolve_book(workbook_id: str | None) -> Any:
    xw_ = _require_xlwings()
    try:
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
    except WorkbookNotFoundError:
        raise
    except Exception as e:
        _translate_excel_error(e)  # macOS 권한 거부면 안내로 바꿔 던진다
        raise
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
    # 병합 감지는 COM 속성(MergeCells/MergeArea/Rows.Count/Row/Column)에 의존한다.
    # macOS는 xlwings가 appscript 참조를 주는데 그 속성 이름이 다르고 동적 참조라
    # `getattr(..., None)` 가드도 안 먹어 안전하지 않다. 지금은 non-Windows에서
    # 병합 감지를 건너뛴다(병합 없음으로 간주) - 일반 셀 읽기/쓰기는 그대로
    # 동작하고, 병합 anchor 리다이렉트만 macOS에서 빠진다(문서화된 제약).
    if platform.system() != "Windows":
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
                # 실측(2026-08-17): 다중 셀 range의 `.formula`는 tuple로 온다
                # (read_range와 같은 이유) - list만 걸러내면 늘 else로 빠져서
                # 셀별 formula를 못 집어낸다.
                formula = (
                    formula_block[r][c]
                    if isinstance(formula_block, (list, tuple))
                    else formula_block
                )
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
    # 실측(2026-08-17, 실제 Excel/xlwings 0.36.16): 다중 셀 range의 `.formula`는
    # list가 아니라 **tuple의 tuple**로 온다(`.value`는 `.options(ndim=2)`로
    # 정규화되지만 `.formula`는 그대로 raw COM 반환값이다). `isinstance(...,
    # list)` 체크는 tuple을 못 걸러서 이미 2차원인 걸 `[[그대로]]`로 한 번 더
    # 감싸버리는 버그가 있었다 - 1x1(스칼라)/다중 셀(tuple) 둘 다 명시적으로
    # 정규화한다.
    raw_formulas = rng.formula
    if isinstance(raw_formulas, (list, tuple)):
        formulas = [list(row) for row in raw_formulas]
    else:
        formulas = [[raw_formulas]]
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


# ---- 시트 관리(목록/추가/이름변경/삭제/이동) ----
# 에이전트가 "시트 하나 만들어" / "시트 옮겨" 같은 요청을 실제로 수행하려면
# 셀 편집만으로는 부족하다 - 시트 자체를 다루는 도구가 있어야 한다.


def list_sheets(workbook_id: str | None = None) -> list[dict[str, Any]]:
    """통합문서의 시트 목록(순서·이름). 0-based index."""
    book = _resolve_book(workbook_id)
    return [{"index": i, "name": s.name} for i, s in enumerate(book.sheets)]


def add_sheet(
    workbook_id: str | None = None,
    name: str | None = None,
    before: int | str | None = None,
    after: int | str | None = None,
) -> str:
    """새 시트를 만든다. before/after로 위치를 줄 수 있고, 안 주면 맨 뒤에
    붙인다(기존 시트 순서를 안 밀도록). 만들어진 시트 이름을 반환한다."""
    book = _resolve_book(workbook_id)
    kwargs: dict[str, Any] = {}
    if name:
        kwargs["name"] = name
    if before is not None:
        kwargs["before"] = book.sheets[before]
    elif after is not None:
        kwargs["after"] = book.sheets[after]
    else:
        n = len(book.sheets)
        if n:
            kwargs["after"] = book.sheets[n - 1]
    try:
        ws = book.sheets.add(**kwargs)
    except Exception as e:
        _translate_excel_error(e)
        raise
    return ws.name


def rename_sheet(workbook_id: str | None, sheet: int | str, new_name: str) -> str:
    """시트 이름을 바꾼다."""
    ws = _sheet(workbook_id, sheet)
    ws.name = new_name
    return new_name


def delete_sheet(workbook_id: str | None, sheet: int | str) -> None:
    """시트를 삭제한다(마지막 한 장은 Excel이 삭제를 막는다)."""
    ws = _sheet(workbook_id, sheet)
    ws.delete()


def move_sheet(
    workbook_id: str | None,
    sheet: int | str,
    before: int | str | None = None,
    after: int | str | None = None,
) -> None:
    """시트를 다른 위치로 이동(재정렬)한다. before 또는 after 중 하나로 대상
    시트를 지정한다. xlwings에 크로스플랫폼 move가 없어 네이티브 객체로 옮긴다
    (Windows COM=`Move(Before/After)`, macOS appscript=`move(to=...)`)."""
    if before is None and after is None:
        raise ValueError("before 또는 after로 이동 위치를 지정하세요.")
    book = _resolve_book(workbook_id)
    ws = book.sheets[sheet]
    target = book.sheets[before] if before is not None else book.sheets[after]
    if platform.system() == "Windows":
        if before is not None:
            ws.api.Move(Before=target.api)
        else:
            ws.api.Move(After=target.api)
        return
    # macOS(appscript): 대상 시트의 앞/뒤를 삽입 위치로 지정해 옮긴다.
    # (appscript 참조의 `.before`/`.after`가 element-relative insertion location)
    try:
        loc = target.api.before if before is not None else target.api.after
        ws.api.move(to=loc)
    except Exception as e:
        _translate_excel_error(e)
        raise


# ---- 서식(채우기 색/굵게/글자색) ----
# "결함조치완료인 행만 노란색으로 색칠" 같은 요청을 하려면 값 편집만으로는
# 안 되고 셀 서식을 바꾸는 도구가 있어야 한다.
_NAMED_COLORS = {
    "yellow": (255, 255, 0), "노랑": (255, 255, 0), "노란색": (255, 255, 0),
    "red": (255, 0, 0), "빨강": (255, 0, 0), "빨간색": (255, 0, 0),
    "green": (0, 176, 80), "초록": (0, 176, 80), "초록색": (0, 176, 80),
    "blue": (0, 112, 192), "파랑": (0, 112, 192), "파란색": (0, 112, 192),
    "orange": (255, 192, 0), "주황": (255, 192, 0),
    "gray": (191, 191, 191), "grey": (191, 191, 191), "회색": (191, 191, 191),
    "white": (255, 255, 255), "흰색": (255, 255, 255),
    "black": (0, 0, 0), "검정": (0, 0, 0),
}


def _rgb(color: Any) -> tuple[int, int, int]:
    """색 입력을 (r,g,b)로 정규화한다. `#RRGGBB`/`RRGGBB` 헥스, `(r,g,b)`,
    또는 흔한 색 이름(yellow/노란색 등)을 받는다."""
    if isinstance(color, (list, tuple)) and len(color) == 3:
        return (int(color[0]), int(color[1]), int(color[2]))
    s = str(color).strip().lower()
    if s in _NAMED_COLORS:
        return _NAMED_COLORS[s]
    h = s.lstrip("#")
    if len(h) == 6:
        try:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except ValueError:
            pass
    raise ValueError(f"색상 형식을 알 수 없습니다: {color!r} (예: '#FFFF00' 또는 'yellow'/'노란색')")


def set_fill_color(
    workbook_id: str | None,
    sheet: int | str,
    row0: int,
    col0: int,
    row1: int,
    col1: int,
    color: Any,
) -> None:
    """직사각형 범위의 배경(채우기) 색을 칠한다. `color`가 None/'none'/''이면
    채우기를 제거한다. 특정 행들만 칠하려면 각 행 범위로 여러 번 호출한다."""
    ws = _sheet(workbook_id, sheet)
    rng = ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1))
    if color is None or str(color).strip().lower() in ("none", ""):
        rng.color = None
    else:
        rng.color = _rgb(color)


def format_range(
    workbook_id: str | None,
    sheet: int | str,
    row0: int,
    col0: int,
    row1: int,
    col1: int,
    fill_color: Any = None,
    bold: bool | None = None,
    font_color: Any = None,
) -> None:
    """범위에 여러 서식을 한 번에 적용한다. 지정한 것만 바꾼다(None은 그대로).
    fill_color='none'이면 채우기 제거."""
    ws = _sheet(workbook_id, sheet)
    rng = ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1))
    if fill_color is not None:
        rng.color = None if str(fill_color).strip().lower() in ("none", "") else _rgb(fill_color)
    if bold is not None:
        rng.font.bold = bool(bold)
    if font_color is not None:
        rng.font.color = _rgb(font_color)


# ---- 구조/서식 일반 기능(숫자서식·병합·정렬·자동맞춤·크기·지우기) ----
# "엑셀에 침투했으면 전체 기능을 다 쓸 수 있어야 한다"는 요구에 맞춰, 셀 값
# 편집을 넘어 실무에서 자주 쓰는 통합문서 조작을 폭넓게 노출한다. 전부 xlwings
# 로 Windows/macOS 공통 동작한다.


def _rng(workbook_id: str | None, sheet: int | str, row0: int, col0: int, row1: int, col1: int) -> Any:
    ws = _sheet(workbook_id, sheet)
    return ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1))


def set_number_format(workbook_id: str | None, sheet: int | str,
                      row0: int, col0: int, row1: int, col1: int, format_code: str) -> None:
    """범위의 표시 형식을 지정한다. 예: '#,##0'(천단위), '0.00%'(백분율),
    'yyyy-mm-dd'(날짜), '₩#,##0'(원화), '@'(텍스트)."""
    _rng(workbook_id, sheet, row0, col0, row1, col1).number_format = format_code


def merge_cells(workbook_id: str | None, sheet: int | str,
                row0: int, col0: int, row1: int, col1: int) -> None:
    """범위를 하나로 병합한다(제목/헤더에 자주 씀)."""
    _rng(workbook_id, sheet, row0, col0, row1, col1).merge()


def unmerge_cells(workbook_id: str | None, sheet: int | str,
                  row0: int, col0: int, row1: int, col1: int) -> None:
    """병합을 해제한다."""
    _rng(workbook_id, sheet, row0, col0, row1, col1).unmerge()


def clear_range(workbook_id: str | None, sheet: int | str,
                row0: int, col0: int, row1: int, col1: int,
                contents: bool = True, formats: bool = False) -> None:
    """범위를 지운다. contents=값/수식, formats=서식. 둘 다 True면 전부 지운다."""
    rng = _rng(workbook_id, sheet, row0, col0, row1, col1)
    if contents and formats:
        rng.clear()
    elif contents:
        rng.clear_contents()
    elif formats:
        try:
            rng.api.ClearFormats()  # Windows COM
        except Exception:
            try:
                rng.api.clear_formats()  # macOS appscript
            except Exception:
                pass


def autofit(workbook_id: str | None, sheet: int | str,
            row0: int | None = None, col0: int | None = None,
            row1: int | None = None, col1: int | None = None) -> None:
    """열 너비/행 높이를 내용에 맞춘다. 범위를 안 주면 사용 범위 전체."""
    ws = _sheet(workbook_id, sheet)
    if row0 is None:
        ws.used_range.autofit()
    else:
        ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1)).autofit()


def set_column_width(workbook_id: str | None, sheet: int | str, col: int, width: float) -> None:
    """열 너비를 지정한다(문자 단위)."""
    _sheet(workbook_id, sheet).range((1, col + 1)).column_width = width


def set_row_height(workbook_id: str | None, sheet: int | str, row: int, height: float) -> None:
    """행 높이를 지정한다(포인트 단위)."""
    _sheet(workbook_id, sheet).range((row + 1, 1)).row_height = height


def sort_range(workbook_id: str | None, sheet: int | str,
               row0: int, col0: int, row1: int, col1: int,
               key_col: int, ascending: bool = True, has_header: bool = True) -> None:
    """범위를 특정 열 기준으로 정렬한다. key_col은 시트 절대 열 인덱스(0-based).
    **Excel 네이티브 정렬**을 써서 행 전체를 서식·수식까지 보존하며 재배치한다
    (값만 다시 써넣어 수식/서식을 날리는 방식이 아니다 - 원본 비파괴)."""
    ws = _sheet(workbook_id, sheet)
    rng = ws.range((row0 + 1, col0 + 1), (row1 + 1, col1 + 1))
    key = ws.range((row0 + 1, key_col + 1), (row1 + 1, key_col + 1))
    if platform.system() == "Windows":
        rng.api.Sort(
            Key1=key.api,
            Order1=1 if ascending else 2,   # 1=xlAscending, 2=xlDescending
            Header=1 if has_header else 2,   # 1=xlYes, 2=xlNo
            Orientation=1,                   # 1=xlSortRows
        )
        return
    # macOS(appscript): Excel의 sort 명령. 행 전체를 보존하며 정렬한다.
    try:
        import appscript  # type: ignore[import-not-found]

        order = appscript.k.sort_ascending if ascending else appscript.k.sort_descending
        header = appscript.k.header_yes if has_header else appscript.k.header_no
        rng.api.sort(key1=key.api, order1=order, header=header)
    except Exception as e:
        _translate_excel_error(e)
        raise
