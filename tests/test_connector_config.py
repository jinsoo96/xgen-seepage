"""config.py를 실제 파일 IO로 검증한다. keyring만 인메모리 가짜로 바꿔서
(진짜 Windows Credential Manager를 테스트에서 건드리지 않기 위해) 토큰
저장/조회/삭제 왕복을 실제로 돈다."""
from __future__ import annotations

from pathlib import Path

import pytest

from xgen_seepage.connector import config as cfgmod


class _FakeKeyring:
    def __init__(self) -> None:
        self._store: dict[tuple[str, str], str] = {}

    def get_password(self, service: str, key: str) -> str | None:
        return self._store.get((service, key))

    def set_password(self, service: str, key: str, value: str) -> None:
        self._store[(service, key)] = value

    def delete_password(self, service: str, key: str) -> None:
        try:
            del self._store[(service, key)]
        except KeyError:
            import keyring.errors

            raise keyring.errors.PasswordDeleteError("not found")


@pytest.fixture
def fake_keyring(monkeypatch: pytest.MonkeyPatch) -> _FakeKeyring:
    fake = _FakeKeyring()
    monkeypatch.setattr(cfgmod.keyring, "get_password", fake.get_password)
    monkeypatch.setattr(cfgmod.keyring, "set_password", fake.set_password)
    monkeypatch.setattr(cfgmod.keyring, "delete_password", fake.delete_password)
    return fake


@pytest.fixture
def isolated_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setattr(cfgmod.Path, "home", staticmethod(lambda: tmp_path))
    return tmp_path


def test_load_config_missing_file_returns_defaults(isolated_home: Path) -> None:
    cfg = cfgmod.load_config()
    assert cfg.server_url == ""
    assert cfg.user_id == ""


def test_save_then_load_roundtrip(isolated_home: Path) -> None:
    cfg = cfgmod.SeepageConfig(server_url="https://xgen.example.com", user_id="u1", username="tester")
    cfgmod.save_config(cfg)

    reloaded = cfgmod.load_config()
    assert reloaded.server_url == "https://xgen.example.com"
    assert reloaded.user_id == "u1"
    assert reloaded.username == "tester"

    on_disk = isolated_home / ".xgen-seepage" / "config.json"
    assert on_disk.exists()
    # 토큰은 이 파일에 절대 평문으로 있으면 안 된다.
    assert "access_token" not in on_disk.read_text(encoding="utf-8")


def test_tokens_roundtrip_via_keyring(fake_keyring: _FakeKeyring) -> None:
    assert cfgmod.get_access_token() is None
    cfgmod.set_tokens("acc-123", "ref-456")
    assert cfgmod.get_access_token() == "acc-123"
    assert cfgmod.get_refresh_token() == "ref-456"

    cfgmod.clear_tokens()
    assert cfgmod.get_access_token() is None
    assert cfgmod.get_refresh_token() is None


def test_clear_tokens_when_nothing_stored_does_not_raise(fake_keyring: _FakeKeyring) -> None:
    cfgmod.clear_tokens()  # 아무것도 없는데 지워도 조용히 넘어가야 한다


def test_set_tokens_raises_clear_error_when_keychain_backend_broken(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # 2026-08-13 실측: PyInstaller로 얼린 exe를 대화형 로그온 세션이 아닌
    # 컨텍스트에서 실행하면 keyring의 Windows 백엔드가 keyring.errors 로 안
    # 감싸인 raw OSError(WinError 1312)를 그대로 던진다. 여기서 그 상황을
    # 흉내내 KeyringUnavailableError로 바뀌는지, 원시 예외가 새지 않는지 확인.
    def broken_set_password(_service: str, _key: str, _value: str) -> None:
        raise OSError("WinError 1312: logon session does not exist")

    monkeypatch.setattr(cfgmod.keyring, "set_password", broken_set_password)

    with pytest.raises(cfgmod.KeyringUnavailableError, match="키체인"):
        cfgmod.set_tokens("acc-123", "ref-456")


def test_get_token_returns_none_when_keychain_backend_broken(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # 읽기 실패는 크래시가 아니라 "토큰 없음"(재로그인 유도)으로 처리돼야
    # run/status 같은 조회 경로가 죽지 않는다.
    def broken_get_password(_service: str, _key: str) -> None:
        raise OSError("WinError 1312: logon session does not exist")

    monkeypatch.setattr(cfgmod.keyring, "get_password", broken_get_password)

    assert cfgmod.get_access_token() is None
    assert cfgmod.get_refresh_token() is None


def test_clear_tokens_does_not_raise_when_keychain_backend_broken(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def broken_delete_password(_service: str, _key: str) -> None:
        raise OSError("WinError 1312: logon session does not exist")

    monkeypatch.setattr(cfgmod.keyring, "delete_password", broken_delete_password)

    cfgmod.clear_tokens()  # 로그아웃/정리 경로는 키체인이 깨져도 조용히 성공해야 한다
