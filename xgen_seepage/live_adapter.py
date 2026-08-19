"""라이브 어댑터: xlwings 기반, "지금 열려 있는" 통합문서를 그 자리에서 편집한다.

이것이 xgen-seepage의 핵심 차별점이다. 디스크의 파일을 열고/고치고/저장하는
배치형 편집과 달리, Claude for Excel / ChatGPT for Excel처럼 "지금 화면에 떠
있는 시트를 실시간으로 편집"하려면 파일 저장을 거치지 않고 **살아 있는 Excel
프로세스**에 직접 붙어야 한다.

두 가지 접근을 비교해 xlwings(로컬 COM/AppleScript 자동화)를 선택했다:
  - Office.js 태스크팬 방식(Claude/ChatGPT for Excel 실제 구현): HTTPS로 호스팅되는
    웹앱이 필요하다. `localhost` 인증서로 오프라인 자체는 가능하지만, 매니페스트
    사이드로딩·인증서 신뢰·SharedRuntime 설정 등 설치 마찰이 크다.
  - xlwings COM 자동화(이 모듈의 접근): `pip install`만으로 끝난다. 웹서버도
    인증서도 매니페스트도 없다. 순수 로컬 프로세스 간 통신이라 인터넷 요구사항이
    원천적으로 없다. 대신 Windows/macOS 전용이고 로컬에 Excel 설치가 필요하다.
자세한 리서치 근거는 저장소 루트 `ARCHITECTURE.md` 참조.

병합 셀 판정과 숫자/텍스트 변환 휴리스틱은 `_cellfmt`의 규칙을 그대로
따른다. 대상이 파일이 아니라 xlwings Range라는 점만 다르다.

⚠️ 성능: 각 셀 접근이 COM 왕복 1회다. 시트 전체를 셀 단위로 훑는 것은 느리다.
그래서 값 읽기는 `Range.value`(2차원 리스트 벌크 반환)를 쓰고, 병합 셀 유무는
`UsedRange.MergeCells`로 먼저 빠르게 판정한 뒤(병합이 아예 없으면 O(1)),
있을 때만 셀 단위로 위치를 확인한다. 그마저도 `_MAX_MERGE_SCAN_CELLS`로
상한을 두고 넘으면 `truncated=True`로 알린다. 크면 range를 좁혀 다시 부르라는
신호다(자세한 근거는 이 파일의 `_merge_map` 참고).
"""
from __future__ import annotations

import platform
import unicodedata
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
        want = unicodedata.normalize("NFC", name)
        for app in xw_.apps:
            if app.pid == pid:
                for book in app.books:
                    if book.name == name or unicodedata.normalize("NFC", book.name) == want:
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
    except Exception:
        pass
    # 이름 조회 실패 시 유니코드 정규화로 재시도: macOS는 한글 시트명을 NFD(자모
    # 분해)로 보관해, 채팅에서 온 NFC 리터럴과 바이트가 달라 직접 조회가 빗나간다.
    if isinstance(sheet, str):
        want = unicodedata.normalize("NFC", sheet)
        for s in book.sheets:
            if unicodedata.normalize("NFC", s.name) == want:
                return s
    raise SheetIndexError(f"sheet {sheet!r} not found in {book.name!r}")


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


def get_workbook_overview(workbook_id: str | None = None, preview_rows: int = 6) -> dict[str, Any]:
    """통합문서 **전체 구조**를 한 번에 반환한다 - 모든 시트의 크기·미리보기·
    수식 존재 여부·병합 정보. 시트가 많은 복잡한 파일을 에이전트가 한 눈에
    파악하고 그 안에서 작업하도록 하기 위한 것. 여기서 전체를 파악한 뒤,
    필요한 시트/범위만 read_range로 자세히 읽으면 된다."""
    book = _resolve_book(workbook_id)
    n = len(book.sheets)
    sheets = [
        get_sheet_schema(workbook_id, sheet=i, preview_rows=preview_rows).to_dict()
        for i in range(n)
    ]
    return {"name": book.name, "sheet_count": n, "sheets": sheets}


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


def color_rows_where(
    workbook_id: str | None,
    sheet: int | str,
    condition_col: int,
    match_value: str,
    color: Any,
    target_col0: int,
    target_col1: int,
    data_start_row: int = 0,
    contains: bool = False,
    clear_non_matching: bool = True,
) -> dict[str, Any]:
    """조건 열(condition_col, 0-based)의 값이 match_value인 데이터 행을 찾아 대상
    열 범위(target_col0..target_col1)를 color로 칠한다. **에이전트가 행 번호를
    손으로 짚지 않고 '조건에 맞는 행'을 한 번에** 처리해, 헤더가 복잡한 실무 표에서
    행이 어긋나거나 일부만 칠해지는 사고를 막는다. clear_non_matching이면 대상 열
    범위의 기존 채우기를 먼저 지우고(정확·빠름) 매칭 행만 칠한다. contains면 부분일치.
    data_start_row(0-based)로 헤더 아래 첫 데이터 행을 지정한다."""
    ws = _sheet(workbook_id, sheet)
    rows, cols = _used_dims(ws)
    if rows <= data_start_row:
        return {"matched_rows": 0, "scanned": 0}
    col_vals = ws.range(
        (data_start_row + 1, condition_col + 1), (rows, condition_col + 1)
    ).options(ndim=2).value
    rgb = _rgb(color)
    if clear_non_matching:
        ws.range((data_start_row + 1, target_col0 + 1), (rows, target_col1 + 1)).color = None
    matched: list[int] = []
    target = str(match_value).strip()
    for i, rv in enumerate(col_vals):
        v = rv[0] if rv else None
        sv = "" if v is None else str(v).strip()
        hit = (target in sv) if contains else (sv == target)
        if hit:
            r = data_start_row + i
            ws.range((r + 1, target_col0 + 1), (r + 1, target_col1 + 1)).color = rgb
            matched.append(r + 1)  # 1-based 행 번호
    return {"matched_rows": len(matched), "scanned": len(col_vals), "rows_1based": matched}


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
    if row0 is None or col0 is None or row1 is None or col1 is None:
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


# ---- 구조 편집(행·열 삽입/삭제·복사·표시) ----
# 전부 원본 비파괴: 값을 다시 써넣지 않고 Excel 네이티브 조작으로 옮기므로
# 기존 수식/서식/병합이 그대로 따라 이동한다.


def _col_letter(idx0: int) -> str:
    """0-based 열 인덱스 → 'A','B',...,'AA' 열 문자."""
    n = idx0 + 1
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def insert_rows(workbook_id: str | None, sheet: int | str, before_row: int, count: int = 1) -> None:
    """before_row(0-based) 위치에 빈 행 count개를 삽입한다. 아래 데이터는
    수식/서식째 밀려 내려간다(원본 비파괴)."""
    ws = _sheet(workbook_id, sheet)
    r = before_row + 1
    ws.range(f"{r}:{r + count - 1}").insert("down")


def delete_rows(workbook_id: str | None, sheet: int | str, row: int, count: int = 1) -> None:
    """row(0-based)부터 count개 행을 삭제한다. 아래 데이터가 위로 당겨진다."""
    ws = _sheet(workbook_id, sheet)
    r = row + 1
    ws.range(f"{r}:{r + count - 1}").delete("up")


def insert_columns(workbook_id: str | None, sheet: int | str, before_col: int, count: int = 1) -> None:
    """before_col(0-based) 위치에 빈 열 count개를 삽입한다. 오른쪽 데이터는
    수식/서식째 밀려난다(원본 비파괴)."""
    ws = _sheet(workbook_id, sheet)
    a, b = _col_letter(before_col), _col_letter(before_col + count - 1)
    ws.range(f"{a}:{b}").insert("right")


def delete_columns(workbook_id: str | None, sheet: int | str, col: int, count: int = 1) -> None:
    """col(0-based)부터 count개 열을 삭제한다. 오른쪽 데이터가 왼쪽으로 당겨진다."""
    ws = _sheet(workbook_id, sheet)
    a, b = _col_letter(col), _col_letter(col + count - 1)
    ws.range(f"{a}:{b}").delete("left")


def copy_range(workbook_id: str | None, sheet: int | str,
               row0: int, col0: int, row1: int, col1: int,
               dest_row: int, dest_col: int, dest_sheet: int | str | None = None) -> None:
    """범위를 다른 위치(같은/다른 시트)로 복사한다. 값·수식·서식이 함께 복사되며
    수식의 상대참조는 붙여넣는 위치에 맞게 자동 조정된다."""
    src = _rng(workbook_id, sheet, row0, col0, row1, col1)
    dws = _sheet(workbook_id, dest_sheet if dest_sheet is not None else sheet)
    src.copy(dws.range((dest_row + 1, dest_col + 1)))


def set_rows_visible(workbook_id: str | None, sheet: int | str,
                     row0: int, row1: int, visible: bool) -> None:
    """row0..row1(0-based, 포함) 행을 숨기거나(visible=False) 다시 보이게 한다."""
    ws = _sheet(workbook_id, sheet)
    rng = ws.range(f"{row0 + 1}:{row1 + 1}")
    if platform.system() == "Windows":
        rng.api.EntireRow.Hidden = (not visible)
    else:
        rng.api.entire_row.hidden.set(not visible)


def set_columns_visible(workbook_id: str | None, sheet: int | str,
                        col0: int, col1: int, visible: bool) -> None:
    """col0..col1(0-based, 포함) 열을 숨기거나(visible=False) 다시 보이게 한다."""
    ws = _sheet(workbook_id, sheet)
    rng = ws.range(f"{_col_letter(col0)}:{_col_letter(col1)}")
    if platform.system() == "Windows":
        rng.api.EntireColumn.Hidden = (not visible)
    else:
        rng.api.entire_column.hidden.set(not visible)


# ---- 수식/이름/재계산 ----


def recalculate(workbook_id: str | None = None) -> None:
    """모든 수식을 강제로 재계산한다(자동계산이 꺼져 있거나 값이 안 갱신될 때)."""
    _resolve_book(workbook_id).app.calculate()


def define_name(workbook_id: str | None, name: str, refers_to: str) -> None:
    """이름 정의를 추가한다. refers_to는 '=시트!$A$1:$A$10' 형식의 참조."""
    _resolve_book(workbook_id).names.add(name, refers_to)


def list_names(workbook_id: str | None = None) -> list[dict[str, Any]]:
    """통합문서에 정의된 이름 목록(이름·참조)을 돌려준다."""
    book = _resolve_book(workbook_id)
    out: list[dict[str, Any]] = []
    for n in book.names:
        try:
            refers = n.refers_to
        except Exception:
            refers = None
        out.append({"name": n.name, "refers_to": refers})
    return out


def delete_name(workbook_id: str | None, name: str) -> None:
    """정의된 이름을 삭제한다."""
    _resolve_book(workbook_id).names[name].delete()


# ---- 테두리/틀고정/찾기·바꾸기/유효성/자동필터 ----

_BORDER_WHICH = {
    "all": ("top", "bottom", "left", "right", "inside_h", "inside_v"),
    "outline": ("top", "bottom", "left", "right"),
    "top": ("top",), "bottom": ("bottom",), "left": ("left",), "right": ("right",),
}


def set_borders(workbook_id: str | None, sheet: int | str,
                row0: int, col0: int, row1: int, col1: int,
                which: str = "all", style: str = "thin") -> None:
    """범위에 테두리를 긋는다. which='all'(격자 전체)/'outline'(바깥만)/개별 변,
    style='thin'/'medium'/'thick'. 검정 실선."""
    rng = _rng(workbook_id, sheet, row0, col0, row1, col1)
    edges = _BORDER_WHICH.get(which, _BORDER_WHICH["all"])
    if platform.system() == "Windows":
        idx = {"left": 7, "top": 8, "bottom": 9, "right": 10, "inside_v": 11, "inside_h": 12}
        weight = {"thin": 2, "medium": -4138, "thick": 4}.get(style, 2)
        for e in edges:
            b = rng.api.Borders(idx[e])
            b.LineStyle = 1  # xlContinuous
            b.Weight = weight
        return
    import appscript  # type: ignore[import-not-found]
    from appscript import k as appk  # type: ignore[import-not-found]
    weight = {
        "thin": appk.border_weight_thin,
        "medium": appk.border_weight_medium,
        "thick": appk.border_weight_thick,
    }.get(style, appk.border_weight_thin)
    edge_k = {
        "top": appk.edge_top, "bottom": appk.edge_bottom,
        "left": appk.edge_left, "right": appk.edge_right,
        "inside_h": appk.inside_horizontal, "inside_v": appk.inside_vertical,
    }
    for e in edges:
        b = rng.api.get_border(which_border=edge_k[e])
        b.line_style.set(appk.continuous)
        b.weight.set(weight)


def freeze_panes(workbook_id: str | None, sheet: int | str, rows: int = 1, cols: int = 0) -> None:
    """상단 rows개 행 / 좌측 cols개 열을 고정한다(스크롤해도 안 움직임).
    rows=0 이고 cols=0 이면 고정을 해제한다. 헤더 1줄 고정은 rows=1, cols=0."""
    ws = _sheet(workbook_id, sheet)
    ws.activate()
    book = _resolve_book(workbook_id)
    off = (rows <= 0 and cols <= 0)
    if platform.system() == "Windows":
        win = book.app.api.ActiveWindow
        if off:
            win.FreezePanes = False
            return
        ws.range((rows + 1, cols + 1)).select()
        win.FreezePanes = True
        return
    win = book.app.api.active_window
    if off:
        win.freeze_panes.set(False)
        return
    ws.range((rows + 1, cols + 1)).select()
    win.freeze_panes.set(True)


def find_replace(workbook_id: str | None, sheet: int | str,
                 find: str, replace: str, match_case: bool = False) -> None:
    """시트의 사용 범위에서 find를 replace로 모두 바꾼다."""
    ws = _sheet(workbook_id, sheet)
    if platform.system() == "Windows":
        ws.api.UsedRange.Replace(What=find, Replacement=replace,
                                 LookAt=2, MatchCase=match_case)  # 2=xlPart
        return
    ws.api.used_range.replace(what=find, replacement=replace, match_case=match_case)


def set_data_validation_list(workbook_id: str | None, sheet: int | str,
                             row0: int, col0: int, row1: int, col1: int,
                             values: list[str]) -> None:
    """범위 셀에 드롭다운(목록) 유효성 검사를 건다. values의 항목만 고를 수 있다."""
    rng = _rng(workbook_id, sheet, row0, col0, row1, col1)
    joined = ",".join(str(v) for v in values)
    if platform.system() == "Windows":
        v = rng.api.Validation
        try:
            v.Delete()
        except Exception:
            pass
        v.Add(Type=3, Formula1=joined)  # 3=xlValidateList
        return
    from appscript import k as appk  # type: ignore[import-not-found]
    try:
        rng.api.validation.delete()
    except Exception:
        pass
    rng.api.validation.modify(type=appk.validate_list, formula1=joined)


def set_autofilter(workbook_id: str | None, sheet: int | str,
                   row0: int, col0: int, row1: int, col1: int, on: bool = True) -> None:
    """범위에 자동필터를 켜거나(on=True) 끈다. **Windows 전용** - macOS Excel은
    appscript로 자동필터 적용이 노출되지 않는다."""
    if platform.system() != "Windows":
        raise RuntimeError(
            "macOS Excel은 자동필터 적용을 appscript로 지원하지 않습니다. "
            "정렬(sort_range)이나 조건부 강조(color_rows_where)로 대체하거나 "
            "Windows에서 실행하세요."
        )
    ws = _sheet(workbook_id, sheet)
    rng = _rng(workbook_id, sheet, row0, col0, row1, col1)
    if on:
        if not ws.api.AutoFilterMode:
            rng.api.AutoFilter()
    else:
        if ws.api.AutoFilterMode:
            ws.api.AutoFilterMode = False


def set_wrap_text(workbook_id: str | None, sheet: int | str,
                  row0: int, col0: int, row1: int, col1: int,
                  wrap: bool = True, autofit_rows: bool = True) -> None:
    """범위 셀의 자동 줄바꿈을 켜거나(wrap=True) 끈다. wrap이면 긴 내용이 셀
    안에서 줄바꿈돼 잘리지 않고 다 보이고, autofit_rows면 행 높이를 내용에
    맞춰 늘려 전체가 보이게 한다(열 너비는 건드리지 않는다)."""
    rng = _rng(workbook_id, sheet, row0, col0, row1, col1)
    if platform.system() == "Windows":
        rng.api.WrapText = bool(wrap)
    else:
        rng.api.wrap_text.set(bool(wrap))
    if wrap and autofit_rows:
        try:
            rng.rows.autofit()  # 행 높이만 맞춘다(열은 그대로 두고 줄바꿈 유지)
        except Exception:
            pass


def get_table_region(workbook_id: str | None, sheet: int | str,
                     row: int, col: int) -> dict[str, Any]:
    """지정한 셀(row,col, 0-based)이 속한 표의 범위를 돌려준다. Excel의 현재
    영역(current region: 빈 행·열로 둘러싸인 연속 블록. Ctrl+A로 잡히는 그
    범위)을 기준으로 한다. 표를 값·테두리·서식으로 그려 놓았을 때 그 표의
    실제 경계를 잡는 데 쓴다. 반환 좌표는 0-based(포함)."""
    ws = _sheet(workbook_id, sheet)
    reg = ws.range((row + 1, col + 1)).current_region
    r0, c0 = reg.row - 1, reg.column - 1
    r1, c1 = reg.last_cell.row - 1, reg.last_cell.column - 1
    return {
        "row0": r0, "col0": c0, "row1": r1, "col1": c1,
        "n_rows": r1 - r0 + 1, "n_cols": c1 - c0 + 1,
    }
