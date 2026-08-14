"""Test: chrome_finder finds Chrome via env var, PATH, and known paths.

Registry lookup is Windows-only and requires actual Chrome install to test
positively; we cover the code path negatively (no key → returns None) and rely
on the other layers for positive coverage.

Run: python test_chrome_finder.py
"""
import os
import shutil
import tempfile
from pathlib import Path
from unittest import mock

import chrome_finder


def _make_fake_chrome(dirpath: Path) -> Path:
    p = dirpath / "chrome.exe"
    p.write_bytes(b"MZ\x90\x00")  # 4-byte "exe" — is_file() is all we check
    return p


def _isolated_env(**overrides):
    """os.environ without SHEIN_CHROME_PATH inherited, plus overrides."""
    env = {k: v for k, v in os.environ.items() if k != "SHEIN_CHROME_PATH"}
    env.update(overrides)
    return env


# ── SHEIN_CHROME_PATH override ───────────────────────────────────────────────

def test_env_override_wins_over_everything():
    tmp = Path(tempfile.mkdtemp())
    try:
        fake = _make_fake_chrome(tmp)
        with mock.patch.dict(os.environ, _isolated_env(SHEIN_CHROME_PATH=str(fake)),
                             clear=True):
            # even if all other layers would find something else, env wins
            with mock.patch.object(chrome_finder, "_from_registry",
                                   return_value=r"C:\other\chrome.exe"):
                with mock.patch.object(chrome_finder, "_from_path",
                                       return_value=r"C:\path\chrome.exe"):
                    assert chrome_finder.find_chrome() == str(fake)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_env_override_ignored_if_file_missing():
    with mock.patch.dict(os.environ,
                         _isolated_env(SHEIN_CHROME_PATH=r"Z:\nope\chrome.exe"),
                         clear=True):
        assert chrome_finder._from_env() is None


def test_env_override_strips_quotes():
    tmp = Path(tempfile.mkdtemp())
    try:
        fake = _make_fake_chrome(tmp)
        quoted = f'"{fake}"'
        with mock.patch.dict(os.environ,
                             _isolated_env(SHEIN_CHROME_PATH=quoted),
                             clear=True):
            assert chrome_finder._from_env() == str(fake)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ── shutil.which fallback ────────────────────────────────────────────────────

def test_from_path_uses_shutil_which():
    with mock.patch("chrome_finder.shutil.which") as which:
        which.side_effect = lambda name: (
            r"C:\found\chrome.exe" if name == "chrome.exe" else None
        )
        assert chrome_finder._from_path() == r"C:\found\chrome.exe"


def test_from_path_returns_none_when_which_finds_nothing():
    with mock.patch("chrome_finder.shutil.which", return_value=None):
        assert chrome_finder._from_path() is None


# ── Hardcoded paths ──────────────────────────────────────────────────────────

def test_from_known_paths_returns_first_existing():
    tmp = Path(tempfile.mkdtemp())
    try:
        fake = _make_fake_chrome(tmp)
        with mock.patch.object(chrome_finder, "_KNOWN_PATHS",
                               [r"Z:\nope\chrome.exe", str(fake), r"C:\other\chrome.exe"]):
            assert chrome_finder._from_known_paths() == str(fake)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_from_known_paths_returns_none_when_all_missing():
    with mock.patch.object(chrome_finder, "_KNOWN_PATHS",
                           [r"Z:\a\chrome.exe", r"Z:\b\chrome.exe"]):
        assert chrome_finder._from_known_paths() is None


# ── Precedence / find_chrome ─────────────────────────────────────────────────

def test_find_chrome_returns_none_when_all_layers_miss():
    with mock.patch.dict(os.environ, _isolated_env(), clear=True):
        with mock.patch.object(chrome_finder, "_from_registry", return_value=None):
            with mock.patch("chrome_finder.shutil.which", return_value=None):
                with mock.patch.object(chrome_finder, "_KNOWN_PATHS", []):
                    assert chrome_finder.find_chrome() is None


def test_registry_lookup_prefers_env_when_both_present():
    """Env override must short-circuit before registry is consulted."""
    tmp = Path(tempfile.mkdtemp())
    try:
        fake = _make_fake_chrome(tmp)
        called = {"registry": False}

        def registry_stub():
            called["registry"] = True
            return r"C:\registry\chrome.exe"

        with mock.patch.dict(os.environ,
                             _isolated_env(SHEIN_CHROME_PATH=str(fake)),
                             clear=True):
            with mock.patch.object(chrome_finder, "_from_registry", registry_stub):
                assert chrome_finder.find_chrome() == str(fake)
                assert called["registry"] is False, "registry should not be consulted"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_searched_locations_lists_all_layers():
    out = chrome_finder.searched_locations()
    assert any("SHEIN_CHROME_PATH" in s for s in out)
    assert any("PATH" in s for s in out)
    assert any("Google\\Chrome" in s or "Google/Chrome" in s for s in out)


if __name__ == "__main__":
    test_env_override_wins_over_everything()
    test_env_override_ignored_if_file_missing()
    test_env_override_strips_quotes()
    test_from_path_uses_shutil_which()
    test_from_path_returns_none_when_which_finds_nothing()
    test_from_known_paths_returns_first_existing()
    test_from_known_paths_returns_none_when_all_missing()
    test_find_chrome_returns_none_when_all_layers_miss()
    test_registry_lookup_prefers_env_when_both_present()
    test_searched_locations_lists_all_layers()
    print("ALL PASS")
