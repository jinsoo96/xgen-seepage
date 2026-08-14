from xgen_seepage.connector.hash import sha256_hex


def test_sha256_hex_matches_known_vector() -> None:
    # 표준 SHA-256("abc"). 알고리즘 자체를 잘못 골랐는지(예: md5) 확인하는
    # 앵커. XGEN 게이트웨이가 기대하는 것과 어긋나면 로그인이 항상 실패한다.
    assert sha256_hex("abc") == "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"


def test_sha256_hex_is_deterministic_and_lowercase_hex() -> None:
    h1 = sha256_hex("hello world")
    h2 = sha256_hex("hello world")
    assert h1 == h2
    assert len(h1) == 64
    assert h1 == h1.lower()


def test_sha256_hex_differs_for_different_input() -> None:
    assert sha256_hex("password1") != sha256_hex("password2")
