"""Tests for setup_wizard's pure helpers (Tkinter UI not tested here)."""
import tempfile
from pathlib import Path

import setup_wizard as sw


def test_write_then_load_round_trip():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        sw.write_config_env({
            "SHEIN_SUBMITTED_DIR": "D:\\a\\b",
            "SHEIN_INPUT_FILENAME": "希音链接 - LU.xlsx",
        }, path=p)
        loaded = sw.load_existing_config(p)
        assert loaded["SHEIN_SUBMITTED_DIR"] == "D:\\a\\b"
        assert loaded["SHEIN_INPUT_FILENAME"] == "希音链接 - LU.xlsx"


def test_load_missing_file_returns_empty():
    assert sw.load_existing_config(Path("/__nonexistent__")) == {}


def test_is_first_run_complete_false_when_missing():
    assert sw.is_first_run_complete(Path("/__nonexistent__")) is False


def test_is_first_run_complete_true_when_submitted_dir_set():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("SHEIN_SUBMITTED_DIR=D:\\x\n", encoding="utf-8")
        assert sw.is_first_run_complete(p) is True


def test_is_first_run_complete_false_when_only_other_keys():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("ANTHROPIC_API_KEY=sk-foo\n", encoding="utf-8")
        assert sw.is_first_run_complete(p) is False


def test_mask_api_key_short_or_empty():
    assert sw.mask_api_key("") == ""
    assert sw.mask_api_key("short") == "short"  # too short to mask meaningfully


def test_mask_api_key_long():
    k = "sk-ant-api03-abcdefghijklmnopqrstuvwxyz"
    m = sw.mask_api_key(k)
    assert m.startswith("sk-ant-a"), m
    assert m.endswith("wxyz"), m
    assert "..." in m


if __name__ == "__main__":
    test_write_then_load_round_trip()
    test_load_missing_file_returns_empty()
    test_is_first_run_complete_false_when_missing()
    test_is_first_run_complete_true_when_submitted_dir_set()
    test_is_first_run_complete_false_when_only_other_keys()
    test_mask_api_key_short_or_empty()
    test_mask_api_key_long()
    print("ALL PASS")
