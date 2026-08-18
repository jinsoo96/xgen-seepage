"""XGEN 서버 접속의 TLS 정책. **폐쇄망에서 가장 중요한 부분**이다.

폐쇄망 XGEN은 보통 내부/사설 CA가 발급한 인증서로 HTTPS를 서비스하고, 그 CA는
사내 정책(GPO 등)으로 각 PC의 OS 신뢰 저장소에 이미 설치돼 있다. 그러면
클라이언트가 검증을 끄지 않고도 그 서버를 그대로 신뢰할 수 있어야 한다.

문제: httpx는 기본으로 `certifi`(고정된 공개 CA 목록)만 보고 **OS 신뢰 저장소를
안 본다**. 그래서 내부 CA가 PC에 깔려 있어도 우리 커넥터는 XGEN 인증서를 거부해
접속이 실패한다 - 유일한 우회가 검증 전면 해제뿐이었다(위험). 그래서 폐쇄망
XGEN에 검증을 끄지 않고 붙을 수 있도록, 여기서는 **OS 신뢰 저장소를 쓰는 TLS
컨텍스트**로 맞춘다. (WebSocket
브릿지는 websockets가 기본으로 `ssl.create_default_context()`를 써 이미 OS 저장소를
보지만, 일관성과 truststore의 견고함을 위해 여기서도 같은 컨텍스트를 넘긴다.)
"""
from __future__ import annotations

import ssl


def default_ssl_context() -> ssl.SSLContext:
    """OS 신뢰 저장소로 검증하는 TLS 컨텍스트(폐쇄망 내부 CA 대응).

    `truststore`가 있으면 그걸 쓴다 - OS(윈도우 SChannel/맥 SecureTransport)에
    검증을 위임해, GPO로 밀어넣은 내부 CA나 사내 MITM 프록시 CA까지 그대로
    신뢰한다(pip이 같은 이유로 채택한 방식). 없으면 stdlib 기본 컨텍스트로
    폴백 - 이것도 Windows/macOS에선 `load_default_certs`로 OS 저장소를 읽어,
    certifi만 보는 httpx 기본보다 폐쇄망에 낫다.
    """
    try:
        import truststore

        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except Exception:
        return ssl.create_default_context()


def relaxed_ssl_context() -> ssl.SSLContext:
    """인증서 체인/호스트명 검증을 모두 끈 SSLContext.

    `allow_private_certificate`가 켜졌을 때만 쓴다 - 내부 CA가 OS에 설치돼 있지
    않아 `default_ssl_context()`로도 안 될 때의 명시적 최후 수단이다. 사용자(또는
    배포 스크립트)가 그 서버를 명시적으로 신뢰한다고 표시한 경우에 한정된다.
    """
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def httpx_verify(allow_private_certificate: bool) -> ssl.SSLContext:
    """httpx.AsyncClient(verify=...)에 그대로 넘길 값. 기본은 OS 신뢰 저장소
    컨텍스트(폐쇄망 내부 CA 대응), allow_private_certificate면 검증 전면 해제."""
    return relaxed_ssl_context() if allow_private_certificate else default_ssl_context()
