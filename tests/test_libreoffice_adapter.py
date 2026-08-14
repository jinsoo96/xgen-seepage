"""libreoffice_adapter.py를 실제 LibreOffice로 검증한다.

live_adapter(xlwings/Excel)의 우아한-실패 테스트와 달리, 이건 실제
애플리케이션으로 왕복까지 돈다 - 이 저장소를 만든 SERVER_JS에 Microsoft
Excel은 없지만 LibreOffice는 실제로 설치돼 있어서 가능하다(2026-08-14
확인). LibreOffice가 없는 환경에서는 전부 스킵한다.
"""
from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from xgen_seepage import libreoffice_adapter as lo  # noqa: E402
from xgen_seepage.base import WorkbookNotFoundError  # noqa: E402

pytestmark = pytest.mark.skipif(
    not lo.is_available(), reason="LibreOffice가 이 머신에 설치돼 있지 않음"
)


@pytest.fixture(scope="session", autouse=True)
def _warm_up_libreoffice():
    """soffice를 세션당 한 번만, 첫 테스트 전에 미리 기동+안정화시킨다.

    실측(2026-08-14): 개별 테스트가 각자 지연 기동을 트리거하게 두면(soffice
    포트가 막 열린 시점과 UNO 서비스가 실제 안정화되는 시점 사이 몇 초의
    간극과 겹쳐) 여러 테스트가 동시에 첫 연결을 시도하며 `DisposedException`이
    반복되는 걸 확인했다. 세션 시작 시 한 번, 넉넉히 기다렸다가 실제 호출로
    확인한 뒤 테스트를 시작한다.
    """
    if not lo.is_available():
        yield
        return
    lo.ensure_running()
    import time

    deadline = time.monotonic() + 25
    last_err = None
    while time.monotonic() < deadline:
        try:
            lo.list_open_documents()
            break
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(1)
    else:
        raise RuntimeError(f"LibreOffice 준비 대기 초과: {last_err}")
    yield
    lo.shutdown_worker()


@pytest.fixture
def sample_xlsx(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["h1", "h2", "h3"])
    ws.append(["v1", "v2", "v3"])
    ws.append(["v4", "v5", "v6"])
    ws.merge_cells("A5:B5")
    ws["A5"] = "merged header"
    p = tmp_path / "sample.xlsx"
    wb.save(p)
    return p


@pytest.fixture
def opened_doc(sample_xlsx: Path):
    info = lo.open_document(str(sample_xlsx))
    doc_id = info["path"]
    yield doc_id
    try:
        lo.close_document(doc_id, save=False)
    except Exception:
        pass


def test_open_and_list(sample_xlsx: Path) -> None:
    info = lo.open_document(str(sample_xlsx))
    assert info["already_open"] is False
    try:
        docs = lo.list_open_documents()
        paths = [d["path"] for d in docs]
        assert info["path"] in paths
        found = next(d for d in docs if d["path"] == info["path"])
        assert found["sheets"] == ["Sheet1"]
    finally:
        lo.close_document(info["path"], save=False)


def test_open_twice_reuses_document(sample_xlsx: Path) -> None:
    first = lo.open_document(str(sample_xlsx))
    try:
        second = lo.open_document(str(sample_xlsx))
        assert second["already_open"] is True
        assert second["path"] == first["path"]
    finally:
        lo.close_document(first["path"], save=False)


def test_get_schema(opened_doc: str) -> None:
    schema = lo.get_sheet_schema(opened_doc, "Sheet1", preview_rows=10)
    assert schema.name == "Sheet1"
    assert schema.rows == 5  # 3 data rows + merged header row (row 4 empty counts too)
    assert schema.cols == 3
    assert schema.preview[0] == ["h1", "h2", "h3"]
    assert schema.preview[1] == ["v1", "v2", "v3"]


def test_get_schema_detects_merge(opened_doc: str) -> None:
    schema = lo.get_sheet_schema(opened_doc, "Sheet1", preview_rows=10)
    assert len(schema.merges) == 1
    m = schema.merges[0]
    assert m.anchor == (4, 0)
    assert m.span == (1, 2)


def test_get_cell(opened_doc: str) -> None:
    cell = lo.get_cell(opened_doc, "Sheet1", 0, 1)
    assert cell.text == "h2"
    assert cell.is_anchor is True


def test_set_cell_and_readback(opened_doc: str) -> None:
    old = lo.set_cell(opened_doc, "Sheet1", 1, 1, "changed")
    assert old == "v2"
    cell = lo.get_cell(opened_doc, "Sheet1", 1, 1)
    assert cell.text == "changed"


def test_set_cell_as_formula(opened_doc: str) -> None:
    lo.set_cell(opened_doc, "Sheet1", 2, 2, "1", allow_merge_redirect=False)
    lo.set_cell(opened_doc, "Sheet1", 3, 0, "=1+2", as_formula=True)
    cell = lo.get_cell(opened_doc, "Sheet1", 3, 0)
    assert cell.formula == "=1+2"
    assert cell.value == 3


def test_set_cell_on_merged_non_anchor_raises(opened_doc: str) -> None:
    with pytest.raises(Exception):  # noqa: B017 - ExcelUnavailableError wraps the UNO-side error
        lo.set_cell(opened_doc, "Sheet1", 4, 1, "x")


def test_read_range(opened_doc: str) -> None:
    data = lo.read_range(opened_doc, "Sheet1", 0, 0, 1, 2)
    assert data["values"] == [["h1", "h2", "h3"], ["v1", "v2", "v3"]]


def test_write_range(opened_doc: str) -> None:
    lo.write_range(opened_doc, "Sheet1", 1, 0, [["x1", "x2", "x3"]])
    cell = lo.get_cell(opened_doc, "Sheet1", 1, 0)
    assert cell.text == "x1"


def test_append_row(opened_doc: str) -> None:
    row = lo.append_row(opened_doc, "Sheet1", ["new1", "new2", "new3"])
    assert row == 5
    cell = lo.get_cell(opened_doc, "Sheet1", 5, 0)
    assert cell.text == "new1"


def test_unknown_document_raises(sample_xlsx: Path) -> None:
    with pytest.raises(WorkbookNotFoundError):
        lo.get_sheet_schema(str(sample_xlsx.with_name("does-not-exist.xlsx")), "Sheet1")


def test_save_persists_to_disk(sample_xlsx: Path) -> None:
    info = lo.open_document(str(sample_xlsx))
    doc_id = info["path"]
    try:
        lo.set_cell(doc_id, "Sheet1", 0, 0, "persisted")
        lo.save(doc_id)
        lo.close_document(doc_id, save=False)
    except Exception:
        lo.close_document(doc_id, save=False)
        raise

    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"]["A1"].value == "persisted"
