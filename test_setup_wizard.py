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


# ── New in 0.3.7: backup dir default + markup validation ────────────────────

def test_default_backup_dir_under_submitted():
    assert sw._default_backup_dir(r"D:\a\b") == r"D:\a\b\_backups"


def test_default_backup_dir_empty_returns_empty():
    assert sw._default_backup_dir("") == ""


def test_validate_markup_accepts_common_values():
    assert sw._validate_markup("1.2") == 1.2
    assert sw._validate_markup("1.5") == 1.5
    assert sw._validate_markup("2.0") == 2.0
    assert sw._validate_markup("  1.2  ") == 1.2   # whitespace tolerant


def test_validate_markup_rejects_junk():
    assert sw._validate_markup("") is None
    assert sw._validate_markup("abc") is None
    assert sw._validate_markup(None) is None
    assert sw._validate_markup("0") is None       # zero disallowed
    assert sw._validate_markup("-1.2") is None    # negative disallowed


def test_round_trip_preserves_new_fields():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        sw.write_config_env({
            "SHEIN_SUBMITTED_DIR": r"D:\input",
            "SHEIN_BACKUP_DIR":    r"D:\backups",
            "SHEIN_EBAY_MARKUP":   "1.2",
            "SHEIN_INPUT_FILENAME": "澳洲希音链接 - ZR.xlsx",
        }, path=p)
        loaded = sw.load_existing_config(p)
        assert loaded["SHEIN_BACKUP_DIR"] == r"D:\backups"
        assert loaded["SHEIN_EBAY_MARKUP"] == "1.2"


if __name__ == "__main__":
    test_write_then_load_round_trip()
    test_load_missing_file_returns_empty()
    test_is_first_run_complete_false_when_missing()
    test_is_first_run_complete_true_when_submitted_dir_set()
    test_is_first_run_complete_false_when_only_other_keys()
    test_mask_api_key_short_or_empty()
    test_mask_api_key_long()
    test_default_backup_dir_under_submitted()
    test_default_backup_dir_empty_returns_empty()
    test_validate_markup_accepts_common_values()
    test_validate_markup_rejects_junk()
    test_round_trip_preserves_new_fields()
    print("ALL PASS")
