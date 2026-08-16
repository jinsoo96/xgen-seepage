"""로컬 태스크팬 HTTPS 서버용 자체서명 인증서.

Office(WebView2)가 태스크팬을 로드하려면 `https://` + 신뢰된 인증서가
필요하다(ARCHITECTURE.md §1). Node의 `office-addin-dev-certs`가 하는 일과
같은 두 단계를 순수 파이썬 + Windows 네이티브 도구로 한다:

  1. 인증서/키 생성 - `cryptography`(이미 `mcp` 의존성 트리에 들어있다.
     `pip install`이 새로 필요 없다)로 순수 파이썬 생성. npm/OpenSSL 불필요.
  2. 신뢰 등록 - `certutil.exe -user -addstore Root`(Windows 기본 내장,
     관리자 권한 불필요). `office-addin-dev-certs`도 Windows에서 내부적으로
     이 도구를 쓴다.

CN/SAN을 `localhost`가 아니라 `127.0.0.1` IP 리터럴로 고정한다. Windows에서
`localhost`가 `::1`을 먼저 반환하는 경우가 있어, 서버가 IPv4로만 바인딩돼
있으면 WebView2가 접속을 못 잡는 조용한 실패로 이어질 수 있다(문서화된
함정, `_uno_worker.py`의 시작 경합 문제와 같은 종류의 "겉보기엔 될 것
같은데 실제로 미묘하게 어긋나는" 부류).
"""
from __future__ import annotations

import datetime
import ipaddress
import platform
import subprocess
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

_CN = "127.0.0.1"
_VALIDITY_DAYS = 730
_RENEW_MARGIN_DAYS = 30


def _cert_dir(config_dir: Path) -> Path:
    d = config_dir / "certs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _cert_paths(config_dir: Path) -> tuple[Path, Path]:
    d = _cert_dir(config_dir)
    return d / "taskpane-cert.pem", d / "taskpane-key.pem"


def _is_still_valid(cert_path: Path) -> bool:
    if not cert_path.exists():
        return False
    try:
        cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
    except (OSError, ValueError):
        return False
    now = datetime.datetime.now(datetime.timezone.utc)
    margin = datetime.timedelta(days=_RENEW_MARGIN_DAYS)
    return cert.not_valid_before_utc <= now and now + margin < cert.not_valid_after_utc


def _generate(cert_path: Path, key_path: Path) -> None:
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, _CN)])
    now = datetime.datetime.now(datetime.timezone.utc)
    san = x509.SubjectAlternativeName(
        [
            x509.DNSName("localhost"),
            x509.IPAddress(ipaddress.ip_address("127.0.0.1")),
            x509.IPAddress(ipaddress.ip_address("::1")),
        ]
    )
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=_VALIDITY_DAYS))
        .add_extension(san, critical=False)
        # 리프(서버) 인증서다 - CA:false. Chrome/Edge는 2019년부터 leaf
        # 인증서에 serverAuth EKU가 없으면 신뢰 저장소에 있어도 거부한다.
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                key_encipherment=True,
                content_commitment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False
        )
        .sign(key, hashes.SHA256())
    )
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    key_path.write_bytes(
        key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )


def _trust_windows(cert_path: Path) -> bool:
    """`certutil -user -addstore Root`로 현재 사용자 신뢰 저장소에 등록한다.
    관리자 권한이 필요 없다(`-user` 스코프). 실패해도 예외를 던지지 않고
    False를 반환한다 - 인증서 자체는 유효하니 브라우저/Office가 최초 1회
    수동으로 신뢰하는 경로는 여전히 열려 있다.

    실측(2026-08-17, 이 프로젝트를 만든 비대화형 자동화 세션): `-addstore
    Root`만 `0x80070032 ERROR_NOT_SUPPORTED`로 실패했다. 같은 세션에서 같은
    인증서로 `-addstore CA`/`-addstore My`는 성공했고(파일 포맷 문제가
    아님을 증명), 서명 검증도 통과했다 - Root 저장소 추가만 콕 집어 막힌다.
    OS 키체인(`config.py`의 `KeyringUnavailableError`)이 대화형 로그온
    세션이 아닌 컨텍스트에서 막히는 것과 같은 부류의 제약으로 보인다(진짜
    로그인한 데스크톱 세션이 필요한 Windows 보안 저장소 계열 API). 일반
    사용자가 설치파일을 더블클릭해서 쓰는 정상적인 대화형 세션에서는 이
    호출이 성공해야 하지만, 이 저장소를 만든 환경 자체가 비대화형이라
    직접 확인은 못 했다 - OS 키체인과 동일한 미검증 항목으로 취급한다."""
    try:
        result = subprocess.run(
            ["certutil.exe", "-user", "-addstore", "Root", str(cert_path)],
            capture_output=True,
            text=True,
            timeout=15,
        )
        return result.returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False


def ensure_dev_certificate(config_dir: Path) -> tuple[Path, Path]:
    """유효한 (cert_path, key_path)를 반환한다. 없거나 곧 만료면 새로
    생성하고, Windows에서는 신뢰 저장소 등록까지 시도한다."""
    cert_path, key_path = _cert_paths(config_dir)
    if _is_still_valid(cert_path) and key_path.exists():
        return cert_path, key_path
    _generate(cert_path, key_path)
    if platform.system() == "Windows":
        _trust_windows(cert_path)
    return cert_path, key_path
