"""특정 서버 하나에만 적용하는 TLS 검증 예외. 사설/자체서명 CA를 쓰는
폐쇄망 XGEN 서버 접속을 허용한다.

xgen-connector `src/main/connection-security.ts`(`shouldAllowPrivateCertificate`,
`xgenWebSocketTlsOptions`)와 같은 목적을 위한 헬퍼다(참고 전용 확인. 로직만
같은 취지로 새로 작성, NOTICE 참조). 다만 그쪽은 Electron이 여러 origin에
접속 가능한 브라우저 셸이라 인증서 예외를 **호스트네임 단위로** 좁혀야 하지만,
xgen-seepage의 HttpClient/ConnectorMcpBridge는 처음부터 설정된 XGEN 서버
하나에만 접속하므로 그 스코핑이 구조적으로 이미 보장된다. 클라이언트 전체의
검증을 끄는 것으로 충분하고 안전하다.
"""
from __future__ import annotations

import ssl


def relaxed_ssl_context() -> ssl.SSLContext:
    """인증서 체인/호스트명 검증을 모두 끈 SSLContext.

    `allow_private_certificate`가 켜졌을 때만 쓴다. 사용자(또는 배포
    스크립트)가 그 서버를 명시적으로 신뢰한다고 표시한 경우에 한정된다.
    """
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def httpx_verify(allow_private_certificate: bool) -> bool | ssl.SSLContext:
    """httpx.AsyncClient(verify=...)에 그대로 넘길 값."""
    return relaxed_ssl_context() if allow_private_certificate else True
