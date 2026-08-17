"""셀 값 <-> 텍스트 변환 휴리스틱.

`maybe_number`/`cell_text`의 판정 규칙은 PlateerLab/document-adapter
(`document_adapter/xlsx_adapter.py`)의 동명 로직을 그대로 이식한 것이다
(NOTICE 참조). 이유: COM으로 셀에 문자열을 그대로 쓰면 Excel의 "타이핑하면
자동으로 숫자가 되는" 동작이 일어나지 않아 금액 셀이 문자로 굳어 SUM 등의
수식이 깨진다. 반면 전화번호·우편번호·사번처럼 선행 0이 있거나 대시가
섞인 값을 숫자로 바꾸면 의미가 훼손된다. 두 문서-편집 도구가 같은
휴리스틱을 쓰면 LLM 입장에서 동작이 예측 가능해진다.
"""
from __future__ import annotations

import datetime
import re
from typing import Any

# 깔끔한 숫자 문자열만 매칭 (정수 / 천단위콤마 / 소수). 전화·ID·날짜(대시) 제외.
_NUM_RE = re.compile(
    r"^-?\d+$|^-?\d{1,3}(?:,\d{3})+$|^-?\d+\.\d+$|^-?\d{1,3}(?:,\d{3})+\.\d+$"
)


def maybe_number(s: str) -> int | float | None:
    """금액 등 깔끔한 숫자 문자열을 int/float로. 아니면 None(문자 유지).

    선행 0 정수(우편번호·사번 등)와 대시 포함(전화·날짜)은 문자로 둔다.
    """
    t = s.strip()
    if not _NUM_RE.match(t):
        return None
    plain = t.replace(",", "")
    intpart = plain.lstrip("-").split(".")[0]
    if len(intpart) > 1 and intpart.startswith("0"):
        return None
    try:
        f = float(plain)
        return int(f) if "." not in plain else f
    except ValueError:
        return None


def cell_text(v: Any) -> str:
    """셀 값을 사람이 읽는 텍스트로. 날짜는 시간이 0이면 날짜만 표시.

    실측(2026-08-17, 실제 Excel/COM): Excel엔 내부 정수 타입이 없어
    `Range.Value`는 42처럼 딱 떨어지는 값도 항상 파이썬 float(42.0)로
    돌아온다. 그대로 `str()`하면 "42.0"이 되어 Excel 화면에 실제로 보이는
    "42"와 어긋난다 - 정수값 float은 트레일링 `.0` 없이 보여준다.
    """
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        if (v.hour, v.minute, v.second, v.microsecond) == (0, 0, 0, 0):
            return v.date().isoformat()
        return v.isoformat(sep=" ")
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def coerce_for_write(value: str) -> Any:
    """set_cell류에 들어온 문자열을 Excel에 쓸 실제 타입으로 변환."""
    num = maybe_number(value)
    return num if num is not None else value
