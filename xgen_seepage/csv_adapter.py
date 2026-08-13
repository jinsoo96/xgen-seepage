"""CSV 어댑터: 인코딩/구분자 자동 감지 + 표 스키마 조회 + 셀 단위 편집.

CSV는 XLSX와 달리 셀 단위 랜덤 액세스 포맷이 아니라서, 편집은 전체를
읽어 메모리에서 고친 뒤 원본 인코딩/구분자를 보존해 통째로 다시 쓰는
방식(read-modify-write-whole-file)을 쓴다. 파일 자체에는 수식 개념이
없으므로 `CellContent.formula`는 항상 None이다 — CSV에서 "수식처럼
계산되는 셀"이 필요하면 `open_in_excel`로 살아 있는 Excel에 올려
live_adapter로 넘어가는 편이 맞다(수식은 애초에 스프레드시트 애플리케이션의
기능이지 CSV 파일 포맷의 기능이 아니다).

인코딩 자동 감지(BOM 우선 → chardet → 후보 목록 순차 시도 → latin-1 최종
폴백)는 PlateerLab/xgen-doc2chunk의 `csv_helper/csv_encoding.py`와 동일한
전략을 그대로 이식했다(NOTICE 참조) — 실사용 데이터(EUC-KR로 저장된 국내
공공/사내 CSV 등)에서 검증된 순서이기 때문에 새로 설계하지 않았다.
"""
from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .base import CellContent, CellOutOfBoundsError, SeepageError, SheetSchema
from ._cellfmt import cell_text

try:
    import chardet
except ImportError:  # pragma: no cover - chardet은 필수 의존성이지만 방어적으로 처리
    chardet = None  # type: ignore[assignment]

# BOM 우선 감지용 시그니처. PlateerLab/xgen-doc2chunk csv_helper/csv_encoding.py 이식.
_BOM_TABLE: list[tuple[bytes, str]] = [
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32-le"),
    (b"\x00\x00\xfe\xff", "utf-32-be"),
    (b"\xff\xfe", "utf-16-le"),
    (b"\xfe\xff", "utf-16-be"),
]

# 국내 데이터에서 흔한 순서: BOM/chardet가 실패했을 때 순차 시도.
_ENCODING_CANDIDATES = ["utf-8", "cp949", "euc-kr", "cp1252"]


class CsvFormatError(SeepageError):
    """CSV 파싱 실패(디코딩 불가·구분자 판별 불가 등)."""


def detect_bom(data: bytes) -> str | None:
    for sig, enc in _BOM_TABLE:
        if data.startswith(sig):
            return enc
    return None


def detect_encoding(data: bytes, preferred_encoding: str | None = None) -> tuple[str, str]:
    """(decoded_text, encoding) 튜플. latin-1 폴백이 있어 항상 성공한다."""
    bom = detect_bom(data)
    if bom:
        try:
            return data.decode(bom), bom
        except UnicodeDecodeError:
            pass

    if preferred_encoding:
        try:
            return data.decode(preferred_encoding), preferred_encoding
        except (UnicodeDecodeError, LookupError):
            pass

    if chardet is not None:
        detected = chardet.detect(data[:10_000])
        enc = detected.get("encoding") if detected else None
        confidence = detected.get("confidence", 0) if detected else 0
        if enc and confidence > 0.7:
            try:
                return data.decode(enc), enc
            except (UnicodeDecodeError, LookupError):
                pass

    for enc in _ENCODING_CANDIDATES:
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue

    return data.decode("latin-1", errors="replace"), "latin-1"


def sniff_dialect(sample: str) -> type[csv.Dialect]:
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.excel  # 기본 comma-dialect 폴백


@dataclass
class CsvTable:
    path: Path
    encoding: str
    dialect: type[csv.Dialect]
    has_header: bool
    header: list[str] | None
    rows: list[list[str]]

    @property
    def n_rows(self) -> int:
        return len(self.rows) + (1 if self.has_header else 0)

    @property
    def n_cols(self) -> int:
        width = max((len(r) for r in self.rows), default=0)
        if self.header:
            width = max(width, len(self.header))
        return width


def load_table(
    path: str | Path,
    *,
    preferred_encoding: str | None = None,
    has_header: bool = True,
) -> CsvTable:
    """CSV를 읽어 CsvTable로 반환한다.

    `has_header` 자동판별은 일부러 하지 않는다 — stdlib `csv.Sniffer().has_header()`는
    첫 행과 데이터 행의 컬럼 "타입"을 비교해 판단하는데, 실사용 CSV는 헤더도
    데이터도 전부 문자열인 열이 흔해서(예: "이름,부서" 헤더 다음에 "홍길동,AI팀")
    헤더가 명백히 있는데도 False로 오판하는 경우가 잦다(이 저장소
    `tests/test_csv_adapter.py::test_get_schema_preview`가 그 실패를 실제로
    재현했다). 업무용 표 데이터는 첫 행이 헤더인 경우가 압도적으로 많으므로
    기본값은 True로 고정하고, 헤더가 없는 파일은 호출 측(에이전트 포함)이
    `has_header=False`를 명시하게 한다 — 조용히 틀리는 것보다 명시적인 편이 낫다.
    """
    p = Path(path)
    raw = p.read_bytes()
    text, encoding = detect_encoding(raw, preferred_encoding)
    dialect = sniff_dialect(text[:4096])

    reader = csv.reader(io.StringIO(text), dialect=dialect)
    all_rows = [row for row in reader if row]  # 완전 빈 행 스킵
    if not all_rows:
        raise CsvFormatError(f"{p}: no rows parsed (empty or unreadable)")

    header = all_rows[0] if has_header else None
    body = all_rows[1:] if has_header else all_rows
    return CsvTable(path=p, encoding=encoding, dialect=dialect, has_header=has_header,
                     header=header, rows=body)


def _save(table: CsvTable) -> None:
    buf = io.StringIO()
    writer = csv.writer(buf, dialect=table.dialect)
    if table.has_header and table.header is not None:
        writer.writerow(table.header)
    writer.writerows(table.rows)
    table.path.write_text(buf.getvalue(), encoding=table.encoding, newline="")


def get_schema(path: str | Path, *, preview_rows: int = 12, max_cell_len: int = 60,
                has_header: bool = True) -> SheetSchema:
    """CSV 파일 하나의 크기·헤더·미리보기. xlsx의 get_sheet_schema와 같은
    반환 형태를 써서(SheetSchema) 도구 스키마를 통일한다. `merges`는 CSV에는
    병합 개념이 없으므로 항상 빈 리스트, `has_formulas`는 항상 False."""
    table = load_table(path, has_header=has_header)
    preview_body = table.rows[:preview_rows]
    preview: list[list[str | None]] = []
    if table.has_header and table.header is not None:
        preview.append([cell_text(v)[:max_cell_len] for v in table.header])
    for row in preview_body:
        preview.append([cell_text(v)[:max_cell_len] for v in row])

    return SheetSchema(
        index=0,
        name=table.path.name,
        rows=table.n_rows,
        cols=table.n_cols,
        preview=preview,
        merges=[],
        has_formulas=False,
        truncated=False,
    )


def _data_row_index(table: CsvTable, row: int) -> int:
    """전체 행 번호(row, 헤더 포함 0-based) → rows 리스트 안의 인덱스."""
    return row - 1 if table.has_header else row


def get_cell(path: str | Path, row: int, col: int, has_header: bool = True) -> CellContent:
    table = load_table(path, has_header=has_header)
    if table.has_header and row == 0:
        line = table.header or []
    else:
        idx = _data_row_index(table, row)
        if idx < 0 or idx >= len(table.rows):
            raise CellOutOfBoundsError(f"row {row} out of bounds (0..{table.n_rows - 1})")
        line = table.rows[idx]
    if col < 0 or col >= len(line):
        raise CellOutOfBoundsError(f"col {col} out of bounds for row {row}")
    text = line[col]
    return CellContent(
        row=row, col=col, value=text, formula=None, text=cell_text(text),
        is_anchor=True, anchor=(row, col), span=(1, 1),
    )


def set_cell(path: str | Path, row: int, col: int, value: str, has_header: bool = True) -> str:
    """셀 하나를 고치고 파일 전체를 원래 인코딩/구분자로 다시 쓴다.
    반환값은 이전 값(되돌리기용)."""
    table = load_table(path, has_header=has_header)
    if table.has_header and row == 0:
        if table.header is None:
            table.header = []
        line = table.header
    else:
        idx = _data_row_index(table, row)
        if idx < 0:
            raise CellOutOfBoundsError(f"row {row} out of bounds")
        while idx >= len(table.rows):
            table.rows.append([])
        line = table.rows[idx]
    while col >= len(line):
        line.append("")
    old = line[col]
    line[col] = str(value)
    _save(table)
    return old


def append_row(path: str | Path, values: list[str], has_header: bool = True) -> int:
    """맨 끝에 행을 추가하고 0-based 전체 행 번호(헤더 포함)를 반환한다."""
    table = load_table(path, has_header=has_header)
    table.rows.append([str(v) for v in values])
    _save(table)
    return table.n_rows - 1


def write_table(
    path: str | Path,
    rows: list[list[str]],
    *,
    header: list[str] | None = None,
    encoding: str = "utf-8-sig",
) -> None:
    """CSV를 새로 만들거나 통째로 덮어쓴다. 기본 인코딩은 utf-8-sig —
    BOM을 붙여야 Excel이 더블클릭으로 열었을 때 한글이 깨지지 않는다."""
    p = Path(path)
    buf = io.StringIO()
    writer = csv.writer(buf)
    if header is not None:
        writer.writerow(header)
    writer.writerows(rows)
    p.write_text(buf.getvalue(), encoding=encoding, newline="")


def open_in_excel(path: str | Path) -> dict[str, Any]:
    """CSV를 로컬 Excel에서 실제로 연다 — 이후로는 live_adapter의 도구
    (get_live_schema/set_live_cell/...)로 이 파일을 마치 xlsx처럼 실시간
    편집할 수 있다. CSV 자체엔 수식이 없지만, Excel이 열면 셀에 `=SUM(...)`
    같은 수식을 자유롭게 추가할 수 있다(단, "다른 이름으로 저장"하지 않는
    한 저장 시 수식이 아니라 계산된 값으로 CSV에 남는다 — CSV 포맷의 근본
    한계이지 이 도구의 제약이 아니다)."""
    from . import live_adapter

    xw_ = live_adapter._require_xlwings()
    p = Path(path).resolve()
    app = xw_.apps.active or xw_.App(visible=True)
    book = app.books.open(str(p))
    return live_adapter.WorkbookInfo(
        workbook_id=live_adapter._workbook_id(app, book),
        name=book.name,
        full_path=book.fullname if book.fullname != book.name else None,
        app_pid=app.pid,
        sheets=[s.name for s in book.sheets],
        active_sheet=book.sheets.active.name if book.sheets.active else None,
        saved=book.saved,
    ).to_dict()
