"""csv_adapter의 수직 슬라이스를 실제로 검증한다: 인코딩 자동 감지(BOM +
euc-kr), 구분자 자동 감지, 스키마/셀 조회, 셀 편집 후 재로딩, 신규 파일
생성, 그리고 tools.call_tool을 통한 dispatcher 경로까지."""
from __future__ import annotations

from pathlib import Path

import pytest

from xgen_seepage import csv_adapter, tools


def test_detect_bom_utf8_sig(tmp_path: Path) -> None:
    p = tmp_path / "bom.csv"
    p.write_bytes("이름,금액\n홍길동,1000\n".encode("utf-8-sig"))
    table = csv_adapter.load_table(p)
    assert table.encoding == "utf-8-sig"
    assert table.header == ["이름", "금액"]
    assert table.rows == [["홍길동", "1000"]]


def test_detect_cp949_no_bom(tmp_path: Path) -> None:
    p = tmp_path / "euckr.csv"
    p.write_bytes("이름,부서\n김진수,AI플랫폼\n".encode("cp949"))
    table = csv_adapter.load_table(p)
    # cp949/euc-kr 둘 다 이 바이트열을 올바르게 디코딩한다 — 어느 쪽으로
    # 감지되든 내용이 깨지지 않는 것이 핵심.
    assert table.header == ["이름", "부서"]
    assert table.rows == [["김진수", "AI플랫폼"]]


def test_semicolon_dialect_sniff(tmp_path: Path) -> None:
    p = tmp_path / "semicolon.csv"
    p.write_text("a;b;c\n1;2;3\n4;5;6\n", encoding="utf-8")
    table = csv_adapter.load_table(p)
    assert table.header == ["a", "b", "c"]
    assert table.rows == [["1", "2", "3"], ["4", "5", "6"]]


def test_get_schema_preview(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1,h2\nv1,v2\nv3,v4\n", encoding="utf-8")
    schema = csv_adapter.get_schema(p, preview_rows=1)
    assert schema.name == "t.csv"
    assert schema.rows == 3  # header + 2 data rows
    assert schema.cols == 2
    assert schema.preview == [["h1", "h2"], ["v1", "v2"]]
    assert schema.merges == []
    assert schema.has_formulas is False


def test_get_cell_header_and_data(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1,h2\nv1,v2\n", encoding="utf-8")
    header_cell = csv_adapter.get_cell(p, 0, 1)
    assert header_cell.text == "h2"
    data_cell = csv_adapter.get_cell(p, 1, 0)
    assert data_cell.text == "v1"


def test_set_cell_roundtrip_preserves_encoding_and_dialect(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_bytes("a;b\n1;2\n".encode("utf-8-sig"))
    old = csv_adapter.set_cell(p, 1, 1, "999")
    assert old == "2"

    reloaded = csv_adapter.load_table(p)
    assert reloaded.rows == [["1", "999"]]
    assert reloaded.dialect.delimiter == ";"
    # utf-8-sig로 다시 저장됐는지 바이트 레벨로 확인 (BOM 보존)
    assert p.read_bytes().startswith(b"\xef\xbb\xbf")


def test_append_row(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1,h2\nv1,v2\n", encoding="utf-8")
    row_idx = csv_adapter.append_row(p, ["v3", "v4"])
    assert row_idx == 2  # header(0) + 기존 1행 다음
    table = csv_adapter.load_table(p)
    assert table.rows[-1] == ["v3", "v4"]


def test_write_table_creates_new_file_with_bom(tmp_path: Path) -> None:
    p = tmp_path / "new.csv"
    csv_adapter.write_table(p, [["1", "2"], ["3", "4"]], header=["x", "y"])
    assert p.exists()
    assert p.read_bytes().startswith(b"\xef\xbb\xbf")
    table = csv_adapter.load_table(p)
    assert table.header == ["x", "y"]
    assert table.rows == [["1", "2"], ["3", "4"]]


def test_out_of_bounds_raises(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1\nv1\n", encoding="utf-8")
    with pytest.raises(csv_adapter.CellOutOfBoundsError):
        csv_adapter.get_cell(p, 0, 5)


def test_tools_dispatch_inspect_and_edit(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1,h2\nv1,v2\n", encoding="utf-8")

    schema = tools.call_tool("inspect_csv", {"path": str(p)})
    assert schema["name"] == "t.csv"
    assert "error" not in schema

    result = tools.call_tool("set_csv_cell", {"path": str(p), "row": 1, "col": 1, "value": "changed"})
    assert result == {"old_value": "v2"}

    cell = tools.call_tool("get_csv_cell", {"path": str(p), "row": 1, "col": 1})
    assert cell["text"] == "changed"


def test_tools_dispatch_unknown_tool() -> None:
    result = tools.call_tool("does_not_exist", {})
    assert result == {"error": "unknown tool: does_not_exist"}


def test_tools_dispatch_wraps_errors(tmp_path: Path) -> None:
    p = tmp_path / "t.csv"
    p.write_text("h1\nv1\n", encoding="utf-8")
    result = tools.call_tool("get_csv_cell", {"path": str(p), "row": 0, "col": 99})
    assert result["error"] == "CellOutOfBoundsError"
