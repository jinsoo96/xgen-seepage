"""XGEN 로그인용 비밀번호 해싱.

XGEN 게이트웨이는 평문이 아니라 **SHA-256 hex 다이제스트**를 그대로
`users.password_hash`와 비교한다. 이 규약을 안 따르면 로그인이 항상
실패한다. PlateerLab/xgen-connector `src/core/hash.ts`가 쓰는 것과 같은
프로토콜(참고 전용 확인, NOTICE 참조). Python 표준 라이브러리 `hashlib`로
새로 구현했다.
"""
from __future__ import annotations

import hashlib


def sha256_hex(plaintext: str) -> str:
    return hashlib.sha256(plaintext.encode("utf-8")).hexdigest()
