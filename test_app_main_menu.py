"""Tests for app_main's menu logic and config check."""
import tempfile
from pathlib import Path

import app_main


def _input_seq(*responses):
    """Returns a callable that yields successive responses."""
    it = iter(responses)
    return lambda prompt="": next(it)


def test_menu_returns_1():
    assert app_main.read_menu_choice(_input_seq("1")) == "1"


def test_menu_returns_2():
    assert app_main.read_menu_choice(_input_seq("2")) == "2"


def test_menu_q_lower_returns_Q():
    assert app_main.read_menu_choice(_input_seq("q")) == "Q"


def test_menu_q_upper_returns_Q():
    assert app_main.read_menu_choice(_input_seq("Q")) == "Q"


def test_menu_empty_returns_Q():
    assert app_main.read_menu_choice(_input_seq("")) == "Q"


def test_menu_invalid_then_valid():
    # First "x" rejected, then "2" accepted
    assert app_main.read_menu_choice(_input_seq("x", "2")) == "2"


def test_has_valid_config_false_when_missing():
    assert app_main.has_valid_config(Path("/__nonexistent__")) is False


def test_has_valid_config_true_when_submitted_dir_set():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("SHEIN_SUBMITTED_DIR=D:\\x\n", encoding="utf-8")
        assert app_main.has_valid_config(p) is True


if __name__ == "__main__":
    test_menu_returns_1()
    test_menu_returns_2()
    test_menu_q_lower_returns_Q()
    test_menu_q_upper_returns_Q()
    test_menu_empty_returns_Q()
    test_menu_invalid_then_valid()
    test_has_valid_config_false_when_missing()
    test_has_valid_config_true_when_submitted_dir_set()
    print("ALL PASS")
