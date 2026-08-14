"""LibreOffice UNO 워커 - LibreOffice가 번들한 파이썬(program\\python.exe)으로만
실행된다. `uno` 모듈은 그 파이썬에 종속된 네이티브 확장이라 일반 venv
파이썬에서 import할 수 없다 - 그래서 `libreoffice_adapter.py`가 이 스크립트를
서브프로세스로 그 파이썬에 태워 실행하고, JSON 파일로 명령/결과를 주고받는다.

의도적으로 표준 라이브러리 + `uno`/`com.sun.star.*`만 쓴다(xgen_seepage 패키지
자체를 import하지 않는다 - LibreOffice 파이썬에는 그 의존성들이 없다).

사용법(상주 모드): python.exe _uno_worker.py --serve
  stdin으로 명령 JSON을 한 줄씩 받아 stdout에 결과 JSON을 한 줄씩 낸다.
  `libreoffice_adapter.py`가 이 프로세스를 한 번만 띄워 계속 재사용한다.

이전엔 호출마다 새 프로세스로 UNO 브릿지를 새로 맺는 방식이었는데, 실측
(2026-08-14)에서 반복되는 연결/해제가 몇 번 지나면 `soffice.bin` CPU가
계속 치솟고 `DisposedException`이 나면서 불안정해지는 걸 확인했다 - 브릿지는
**한 번만 맺고 재사용**하는 게 맞다.
"""
from __future__ import annotations

import json
import sys

sys.path.append(r"C:\Program Files\LibreOffice\program")
sys.path.append(r"C:\Program Files (x86)\LibreOffice\program")

import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402

_HOST = "localhost"
_PORT = 2002

_desktop_cache = None


def _prop(name: str, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _connect():
    """UNO 브릿지는 이 워커 프로세스 생애주기 동안 한 번만 맺고 재사용한다.

    soffice 프로세스가 막 뜬 직후엔 리스닝 소켓은 열려 있어도 내부 서비스
    레지스트리가 아직 준비 안 돼 `DisposedException`이 날 수 있다(2026-08-14
    실측 - TCP 포트가 열린 것과 UNO 서비스가 실제로 응답 가능한 것 사이에
    간극이 있음). 몇 차례 짧게 재시도한다.
    """
    global _desktop_cache
    if _desktop_cache is not None:
        return _desktop_cache
    import time

    last_err = None
    # 실측(2026-08-14): 포트가 열린 뒤에도 soffice 내부 서비스가 완전히
    # 안정화되기까지 8초 넘게 걸리는 걸 확인했다(3초는 부족, 8초는 충분).
    # 40회x0.5초=20초까지 여유를 둔다.
    for attempt in range(40):
        try:
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )
            ctx = resolver.resolve(
                f"uno:socket,host={_HOST},port={_PORT};urp;StarOffice.ComponentContext"
            )
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            # 실제로 서비스가 응답하는지 가벼운 호출로 확인(여기서도 disposed면 재시도).
            desktop.getComponents()
            _desktop_cache = desktop
            return desktop
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(0.5)
    raise RuntimeError(f"UNO 브릿지 연결 실패(재시도 소진): {last_err}")


def _url_to_path(url: str) -> str:
    return uno.fileUrlToSystemPath(url)


def _path_to_url(path: str) -> str:
    return uno.systemPathToFileUrl(path)


def _iter_spreadsheet_docs(desktop):
    comps = desktop.getComponents()
    it = comps.createEnumeration()
    while it.hasMoreElements():
        doc = it.nextElement()
        if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            yield doc


def _find_doc(desktop, doc_path: str | None):
    docs = list(_iter_spreadsheet_docs(desktop))
    if doc_path is None:
        if not docs:
            raise RuntimeError("no spreadsheet document is open")
        return docs[0]
    target = doc_path.replace("/", "\\").lower()
    for doc in docs:
        try:
            p = _url_to_path(doc.getURL()).lower()
        except Exception:
            continue
        if p == target:
            return doc
    raise RuntimeError(f"document not open: {doc_path!r}")


def _used_dims(sheet) -> tuple[int, int]:
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    addr = cursor.RangeAddress
    return addr.EndRow + 1, addr.EndColumn + 1


def _merge_anchor_span(sheet, row: int, col: int):
    """(is_anchor, anchor(row,col), span(rows,cols)).

    `cell.getIsMerged()`에 기대지 않는다 - 실측(2026-08-14, openpyxl로 만든
    xlsx)에서 병합의 **앵커** 셀은 getIsMerged()==True를 정확히 주지만, 같은
    병합의 **non-anchor** 셀에서 호출하면 False가 나오는 비대칭을 확인했다
    (xlsx 저장 포맷과 LibreOffice의 셀별 IsMerged 판정 사이의 상호운용성
    문제로 보임). 대신 커서를 항상 `collapseToMergedArea()`로 접어서 실제
    병합 범위를 직접 얻는다 - 병합이 없는 셀이면 이 호출이 사실상 no-op으로
    자기 자신 1x1을 돌려주므로 모든 셀에 안전하게 쓸 수 있다.
    """
    cell = sheet.getCellByPosition(col, row)
    cursor = sheet.createCursorByRange(cell)
    cursor.collapseToMergedArea()
    addr = cursor.RangeAddress
    anchor = (addr.StartRow, addr.StartColumn)
    span = (addr.EndRow - addr.StartRow + 1, addr.EndColumn - addr.StartColumn + 1)
    return anchor == (row, col), anchor, span


def _cell_formula_or_none(cell) -> str | None:
    formula = cell.getFormula()
    if isinstance(formula, str) and formula.startswith("="):
        return formula
    return None


def _cell_value(cell):
    from com.sun.star.table.CellContentType import EMPTY, TEXT, VALUE, FORMULA

    t = cell.getType()
    if t == EMPTY:
        return None
    if t == TEXT:
        return cell.getString()
    if t in (VALUE, FORMULA):
        return cell.getValue()
    return cell.getString()


def op_open_document(args):
    desktop = _connect()
    path = args["path"]
    url = _path_to_url(path)
    for doc in _iter_spreadsheet_docs(desktop):
        if _url_to_path(doc.getURL()).lower() == path.replace("/", "\\").lower():
            return {"path": path, "already_open": True}
    hidden = bool(args.get("hidden", False))
    props = (_prop("Hidden", hidden),)
    doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
    return {"path": _url_to_path(doc.getURL()), "already_open": False}


def op_close_document(args):
    desktop = _connect()
    doc = _find_doc(desktop, args.get("doc"))
    save = bool(args.get("save", False))
    if save:
        doc.store()
    doc.close(False)
    return {"closed": True}


def op_list_documents(_args):
    desktop = _connect()
    out = []
    for doc in _iter_spreadsheet_docs(desktop):
        sheets = doc.getSheets()
        active_name = None
        try:
            controller = doc.getCurrentController()
            active_name = controller.getActiveSheet().getName()
        except Exception:
            pass
        out.append(
            {
                "path": _url_to_path(doc.getURL()),
                "sheets": list(sheets.getElementNames()),
                "active_sheet": active_name,
                "modified": bool(doc.isModified()),
            }
        )
    return out


def op_get_schema(args):
    desktop = _connect()
    doc = _find_doc(desktop, args.get("doc"))
    sheets = doc.getSheets()
    sheet_ref = args["sheet"]
    sheet = sheets.getByIndex(sheet_ref) if isinstance(sheet_ref, int) else sheets.getByName(sheet_ref)
    rows, cols = _used_dims(sheet)
    preview_rows = int(args.get("preview_rows", 12))
    max_len = int(args.get("max_cell_len", 60))
    visible = min(rows, preview_rows)

    merges = {}
    covered = set()
    preview = [[None] * cols for _ in range(visible)]
    has_formulas = False
    for r in range(rows):
        for c in range(cols):
            cell = sheet.getCellByPosition(c, r)
            is_anchor, anchor, span = _merge_anchor_span(sheet, r, c)
            if span != (1, 1):
                if is_anchor:
                    merges[anchor] = span
                else:
                    covered.add((r, c))
            if not has_formulas:
                f = _cell_formula_or_none(cell)
                if f:
                    has_formulas = True
            if r < visible:
                if (r, c) in covered:
                    continue
                text = cell.getString()
                preview[r][c] = text[:max_len]

    sheet_index = None
    names = list(sheets.getElementNames())
    if isinstance(sheet_ref, int):
        sheet_index = sheet_ref
    else:
        sheet_index = names.index(sheet_ref)

    return {
        "index": sheet_index,
        "name": sheet.getName(),
        "rows": rows,
        "cols": cols,
        "preview": preview,
        "merges": [{"anchor": list(a), "span": list(s)} for a, s in merges.items()],
        "has_formulas": has_formulas,
        "truncated": False,
    }


def _resolve_sheet(desktop, args):
    doc = _find_doc(desktop, args.get("doc"))
    sheets = doc.getSheets()
    sheet_ref = args["sheet"]
    sheet = sheets.getByIndex(sheet_ref) if isinstance(sheet_ref, int) else sheets.getByName(sheet_ref)
    return doc, sheet


def op_get_cell(args):
    desktop = _connect()
    _doc, sheet = _resolve_sheet(desktop, args)
    row, col = int(args["row"]), int(args["col"])
    is_anchor, anchor, span = _merge_anchor_span(sheet, row, col)
    cell = sheet.getCellByPosition(anchor[1], anchor[0])
    value = _cell_value(cell)
    formula = _cell_formula_or_none(cell)
    text = cell.getString()
    return {
        "row": row,
        "col": col,
        "value": value,
        "formula": formula,
        "text": text,
        "is_anchor": is_anchor,
        "anchor": list(anchor),
        "span": list(span),
    }


def op_set_cell(args):
    desktop = _connect()
    _doc, sheet = _resolve_sheet(desktop, args)
    row, col = int(args["row"]), int(args["col"])
    is_anchor, anchor, _span = _merge_anchor_span(sheet, row, col)
    if not is_anchor and not args.get("allow_merge_redirect"):
        raise RuntimeError(
            f"cell ({row},{col}) is part of a merge anchored at "
            f"({anchor[0]},{anchor[1]}). Write to the anchor, or pass allow_merge_redirect."
        )
    cell = sheet.getCellByPosition(anchor[1], anchor[0])
    old_text = cell.getString()
    value = str(args["value"])
    if args.get("as_formula"):
        if not value.startswith("="):
            raise RuntimeError("as_formula=True requires a value starting with '='")
        cell.setFormula(value)
    else:
        cell.setString(value)
        num = _maybe_number(value)
        if num is not None:
            cell.setValue(num)
    return {"old_value": old_text}


def _maybe_number(s: str):
    import re

    m = re.match(r"^-?\d+$|^-?\d+\.\d+$", s.strip())
    if not m:
        return None
    try:
        f = float(s)
        return int(f) if "." not in s else f
    except ValueError:
        return None


def op_read_range(args):
    desktop = _connect()
    _doc, sheet = _resolve_sheet(desktop, args)
    row0, col0, row1, col1 = (int(args[k]) for k in ("row0", "col0", "row1", "col1"))
    n_cells = (row1 - row0 + 1) * (col1 - col0 + 1)
    if n_cells > 50_000:
        raise RuntimeError(f"range too large: {n_cells} cells (max 50000)")
    values, formulas = [], []
    for r in range(row0, row1 + 1):
        vrow, frow = [], []
        for c in range(col0, col1 + 1):
            cell = sheet.getCellByPosition(c, r)
            vrow.append(_cell_value(cell))
            frow.append(_cell_formula_or_none(cell) or cell.getString())
        values.append(vrow)
        formulas.append(frow)
    return {"row0": row0, "col0": col0, "row1": row1, "col1": col1, "values": values, "formulas": formulas}


def op_write_range(args):
    desktop = _connect()
    _doc, sheet = _resolve_sheet(desktop, args)
    row0, col0 = int(args["row0"]), int(args["col0"])
    rows = args["rows"]
    n_cols = max(len(r) for r in rows) if rows else 0
    n_cells = len(rows) * n_cols
    if n_cells > 50_000:
        raise RuntimeError(f"write too large: {n_cells} cells (max 50000)")
    for i, row_vals in enumerate(rows):
        for j, v in enumerate(row_vals):
            cell = sheet.getCellByPosition(col0 + j, row0 + i)
            s = str(v)
            if s.startswith("="):
                cell.setFormula(s)
            else:
                cell.setString(s)
                num = _maybe_number(s)
                if num is not None:
                    cell.setValue(num)
    return {"written_rows": len(rows)}


def op_append_row(args):
    desktop = _connect()
    _doc, sheet = _resolve_sheet(desktop, args)
    rows, _cols = _used_dims(sheet)
    values = args["values"]
    for j, v in enumerate(values):
        if v == "":
            continue
        cell = sheet.getCellByPosition(j, rows)
        s = str(v)
        cell.setString(s)
        num = _maybe_number(s)
        if num is not None:
            cell.setValue(num)
    return {"row": rows}


def op_save(args):
    desktop = _connect()
    doc = _find_doc(desktop, args.get("doc"))
    doc.store()
    return {"saved": True}


_OPS = {
    "open_document": op_open_document,
    "close_document": op_close_document,
    "list_documents": op_list_documents,
    "get_schema": op_get_schema,
    "get_cell": op_get_cell,
    "set_cell": op_set_cell,
    "read_range": op_read_range,
    "write_range": op_write_range,
    "append_row": op_append_row,
    "save": op_save,
}


def _dispatch(command: dict) -> dict:
    global _desktop_cache
    op = command.get("op") or ""
    handler = _OPS.get(op)
    if handler is None:
        return {"ok": False, "error": f"unknown op: {op}"}
    args = command.get("args", {})
    try:
        data = handler(args)
        return {"ok": True, "data": data}
    except Exception as e:  # noqa: BLE001
        # soffice가 막 뜬 직후엔 한동안 연결이 조용히 disposed될 수 있다(실측,
        # 2026-08-14) - 캐시를 비우고 딱 한 번 통째로 재시도한다(재연결
        # 포함). 그래도 안 되면 진짜 에러로 취급.
        if "disposed" in str(e).lower() and _desktop_cache is not None:
            _desktop_cache = None
            try:
                data = handler(args)
                return {"ok": True, "data": data}
            except Exception as e2:  # noqa: BLE001
                return {"ok": False, "error": f"{type(e2).__name__}: {e2}"}
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}


def serve() -> None:
    """stdin에서 명령 JSON을 한 줄씩 읽어 stdout에 결과 JSON을 한 줄씩 쓴다.
    UNO 브릿지는 첫 명령에서 한 번만 맺고 프로세스가 사는 동안 재사용한다."""
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            command = json.loads(line)
        except json.JSONDecodeError as e:
            print(json.dumps({"ok": False, "error": f"bad JSON: {e}"}), flush=True)
            continue
        result = _dispatch(command)
        print(json.dumps(result, ensure_ascii=False), flush=True)


def main() -> None:
    if len(sys.argv) >= 2 and sys.argv[1] == "--serve":
        serve()
        return
    # 하위 호환: 파일 기반 단발 호출(디버깅용).
    cmd_path, result_path = sys.argv[1], sys.argv[2]
    with open(cmd_path, "r", encoding="utf-8") as f:
        command = json.load(f)
    result = _dispatch(command)
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)


if __name__ == "__main__":
    main()
