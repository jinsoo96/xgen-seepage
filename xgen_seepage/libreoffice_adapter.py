"""LibreOffice(UNO) 백엔드 - Microsoft Excel이 없는 환경(폐쇄망 등 Office
라이선스를 못 구하는 경우가 흔하다)에서도 "지금 열려 있는 통합문서를
실시간으로 편집"하는 `live_adapter.py`와 같은 경험을 제공한다.

실측(2026-08-14, SERVER_JS): 이 머신엔 Microsoft Excel이 설치돼 있지 않지만
(COM 클래스 미등록, 3가지 방법으로 확인) LibreOffice는 이미 설치돼 있었다.
실제 열려 있는 xlsx 문서에 UNO로 붙어 셀을 쓰고 즉시 읽어 확인하는 왕복까지
실제로 검증했다. `tests/test_libreoffice_adapter.py`가 이 머신에서 그 왕복을
그대로 재현한다(Excel 쪽 xlwings 테스트와 달리, 이건 실제 애플리케이션으로
돈다 - 우아한 실패 경로만 테스트하는 게 아니다).

아키텍처: `uno` 파이썬 모듈은 LibreOffice가 번들한 파이썬 전용 네이티브
확장이라 이 패키지가 도는 일반 venv/얼린 exe의 파이썬에서는 import할 수
없다. 그래서 실제 UNO 호출은 `_uno_worker.py --serve`를 LibreOffice의
`program\\python.exe`로 **상주 서브프로세스**로 띄우고, stdin/stdout으로
명령/결과 JSON을 한 줄씩 주고받는다. 호출마다 새 프로세스를 띄우는 방식을
먼저 시도했으나(파일 기반 1회성 호출), 실측(2026-08-14)에서 반복되는 UNO
브릿지 연결/해제가 몇 차례 지나자 `soffice.bin`의 CPU가 계속 치솟고
`com.sun.star.lang.DisposedException`이 나면서 불안정해지는 걸 확인했다.
브릿지는 워커 프로세스 생애주기 동안 한 번만 맺고 재사용하는 걸로 고쳤다.
"""
from __future__ import annotations

import json
import queue
import socket
import subprocess
import threading
import time
from pathlib import Path
from typing import Any

from .base import (
    CellContent,
    ExcelUnavailableError,
    MergeInfo,
    SheetSchema,
    WorkbookNotFoundError,
)

_HOST = "localhost"
_PORT = 2002

_worker_proc: subprocess.Popen | None = None
# RLock: _run_worker가 락을 쥔 채로 타임아웃 시 shutdown_worker를 부르므로
# (같은 스레드의 재진입) 일반 Lock이면 데드락난다.
_worker_lock = threading.RLock()

_CANDIDATE_INSTALL_DIRS = [
    r"C:\Program Files\LibreOffice",
    r"C:\Program Files (x86)\LibreOffice",
]


def _find_install_dir() -> Path | None:
    for d in _CANDIDATE_INSTALL_DIRS:
        p = Path(d)
        if (p / "program" / "soffice.exe").exists():
            return p
    return None


def _find_python_exe() -> Path | None:
    install = _find_install_dir()
    if install is None:
        return None
    py = install / "program" / "python.exe"
    return py if py.exists() else None


def is_available() -> bool:
    return _find_install_dir() is not None


def _port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex((_HOST, port)) == 0


def ensure_running(timeout: float = 20.0) -> None:
    """UNO 리스너가 떠 있지 않으면 헤드리스로 기동한다."""
    if _port_open(_PORT):
        return
    install = _find_install_dir()
    if install is None:
        raise ExcelUnavailableError(
            "LibreOffice가 설치돼 있지 않습니다. Excel도 LibreOffice도 없어 "
            "라이브 편집 백엔드를 쓸 수 없습니다."
        )
    soffice = install / "program" / "soffice.exe"
    accept = f"socket,host={_HOST},port={_PORT};urp;StarOffice.ComponentContext"
    subprocess.Popen(
        [
            str(soffice),
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--norestore",
            "--nofirststartwizard",
            "--nologo",
            f"--accept={accept}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if _port_open(_PORT):
            return
        time.sleep(0.3)
    raise ExcelUnavailableError("LibreOffice UNO 리스너가 제한 시간 내에 뜨지 않았습니다.")


def _start_worker() -> subprocess.Popen:
    python_exe = _find_python_exe()
    if python_exe is None:
        raise ExcelUnavailableError("LibreOffice의 번들 파이썬(program\\python.exe)을 찾을 수 없습니다.")
    worker = Path(__file__).with_name("_uno_worker.py")
    return subprocess.Popen(
        [str(python_exe), str(worker), "--serve"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,
    )


def _get_worker() -> subprocess.Popen:
    global _worker_proc
    if _worker_proc is not None and _worker_proc.poll() is None:
        return _worker_proc
    ensure_running()
    _worker_proc = _start_worker()
    return _worker_proc


def shutdown_worker() -> None:
    """상주 워커 프로세스를 종료한다(soffice 자체는 그대로 둔다). 프로세스
    종료/테스트 정리용 - 평소엔 부를 필요 없다."""
    global _worker_proc
    with _worker_lock:
        if _worker_proc is not None:
            try:
                _worker_proc.terminate()
                _worker_proc.wait(timeout=5)
            except Exception:
                try:
                    _worker_proc.kill()
                except Exception:
                    pass
            _worker_proc = None


def _readline_with_timeout(proc: subprocess.Popen, timeout: float) -> str:
    """`readline()`은 자체 타임아웃이 없다 - UNO 쪽에서 조용히 멎는 호출
    (예: 헤드리스에서 뜰 수 없는 숨은 대화상자를 기다리는 경우)이 있으면
    프로세스 전체가 영원히 걸린다(실측, 2026-08-14). 별도 스레드로 읽고
    큐로 시간제한을 건다."""
    q: "queue.Queue[str]" = queue.Queue(maxsize=1)

    def _reader() -> None:
        try:
            q.put(proc.stdout.readline())  # type: ignore[union-attr]
        except Exception:
            q.put("")

    t = threading.Thread(target=_reader, daemon=True)
    t.start()
    try:
        return q.get(timeout=timeout)
    except queue.Empty:
        raise TimeoutError(f"UNO 워커 응답이 {timeout}초 안에 오지 않았습니다.") from None


def _run_worker(op: str, args: dict[str, Any], *, timeout: float = 30.0) -> Any:
    with _worker_lock:
        proc = _get_worker()
        assert proc.stdin is not None and proc.stdout is not None
        try:
            proc.stdin.write(json.dumps({"op": op, "args": args}, ensure_ascii=False) + "\n")
            proc.stdin.flush()
            line = _readline_with_timeout(proc, timeout)
        except TimeoutError:
            # 응답 없는 워커는 신뢰할 수 없다 - 죽이고 다음 호출이 새로 띄우게 한다.
            shutdown_worker()
            raise ExcelUnavailableError(
                f"UNO 작업이 {timeout}초 안에 끝나지 않아 워커를 재시작했습니다 "
                f"(op={op}). 다시 시도해 보세요."
            ) from None
        except (BrokenPipeError, OSError) as e:
            raise ExcelUnavailableError(f"UNO 워커 프로세스와 통신할 수 없습니다: {e}") from e
        if not line:
            stderr = proc.stderr.read() if proc.stderr else ""
            raise ExcelUnavailableError(f"UNO 워커가 응답 없이 종료됐습니다: {stderr[:500]}")
        result = json.loads(line)

    if not result.get("ok"):
        error = result.get("error", "unknown UNO worker error")
        if "not open" in error or "no spreadsheet document" in error:
            raise WorkbookNotFoundError(error)
        raise ExcelUnavailableError(error)
    return result["data"]


def open_document(path: str, *, hidden: bool = True) -> dict[str, Any]:
    """파일을 LibreOffice로 연다(이미 열려 있으면 그대로 재사용). 반환된
    `path`를 이후 도구 호출의 `doc`(workbook_id)로 쓴다."""
    return _run_worker("open_document", {"path": str(path), "hidden": hidden})


def close_document(doc: str | None, *, save: bool = False) -> None:
    _run_worker("close_document", {"doc": doc, "save": save})


def list_open_documents() -> list[dict[str, Any]]:
    """열려 있는 LibreOffice Calc 문서 목록. 각 항목의 `path`가 workbook_id다."""
    return _run_worker("list_documents", {})


def get_sheet_schema(
    doc: str | None,
    sheet: int | str,
    *,
    preview_rows: int = 12,
    max_cell_len: int = 60,
) -> SheetSchema:
    data = _run_worker(
        "get_schema",
        {"doc": doc, "sheet": sheet, "preview_rows": preview_rows, "max_cell_len": max_cell_len},
    )
    return SheetSchema(
        index=data["index"],
        name=data["name"],
        rows=data["rows"],
        cols=data["cols"],
        preview=data["preview"],
        merges=[MergeInfo(anchor=tuple(m["anchor"]), span=tuple(m["span"])) for m in data["merges"]],
        has_formulas=data["has_formulas"],
        truncated=data["truncated"],
    )


def get_cell(doc: str | None, sheet: int | str, row: int, col: int) -> CellContent:
    data = _run_worker("get_cell", {"doc": doc, "sheet": sheet, "row": row, "col": col})
    return CellContent(
        row=data["row"],
        col=data["col"],
        value=data["value"],
        formula=data["formula"],
        text=data["text"],
        is_anchor=data["is_anchor"],
        anchor=tuple(data["anchor"]),
        span=tuple(data["span"]),
    )


def set_cell(
    doc: str | None,
    sheet: int | str,
    row: int,
    col: int,
    value: str,
    *,
    as_formula: bool = False,
    allow_merge_redirect: bool = False,
) -> str:
    data = _run_worker(
        "set_cell",
        {
            "doc": doc,
            "sheet": sheet,
            "row": row,
            "col": col,
            "value": value,
            "as_formula": as_formula,
            "allow_merge_redirect": allow_merge_redirect,
        },
    )
    return data["old_value"]


def read_range(doc: str | None, sheet: int | str, row0: int, col0: int, row1: int, col1: int) -> dict[str, Any]:
    return _run_worker(
        "read_range", {"doc": doc, "sheet": sheet, "row0": row0, "col0": col0, "row1": row1, "col1": col1}
    )


def write_range(doc: str | None, sheet: int | str, row0: int, col0: int, rows: list[list[str]]) -> None:
    _run_worker("write_range", {"doc": doc, "sheet": sheet, "row0": row0, "col0": col0, "rows": rows})


def append_row(doc: str | None, sheet: int | str, values: list[str]) -> int:
    data = _run_worker("append_row", {"doc": doc, "sheet": sheet, "values": values})
    return data["row"]


def save(doc: str | None) -> None:
    _run_worker("save", {"doc": doc})
