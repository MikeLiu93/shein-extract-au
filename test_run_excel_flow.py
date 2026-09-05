"""Tests for the single-workbook run_excel flow:
- _read_pending_rows implicit trigger (B has URL, F+G empty)
- _backup_workbook + _prune_backups
- _archive_legacy_enriched one-shot migration
"""
import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

import run_excel as re_mod
from run_excel import (
    COL_SEQ, COL_URL, COL_PRICE, COL_SHIPPING, COL_VARIANT_FILTER,
    COL_DATE, COL_STATUS,
    EXPECTED_HEADERS,
    _read_pending_rows, _backup_workbook, _prune_backups,
    _archive_legacy_enriched,
)


def _make_template_ws(wb, rows: list[tuple]):
    """rows: sequence of (seq, url, price, shipping, variant, date, status) tuples.
    Use "" or None for empty. First row is header."""
    ws = wb.active
    for ci, h in enumerate(EXPECTED_HEADERS, 1):
        ws.cell(1, ci).value = h
    for ri, row_vals in enumerate(rows, 2):
        for ci, v in enumerate(row_vals, 1):
            ws.cell(ri, ci).value = v
    return ws


# ── _read_pending_rows ──────────────────────────────────────────────────────

def test_pending_selects_only_rows_with_url_and_empty_f_g():
    wb = Workbook()
    ws = _make_template_ws(wb, [
        (1, "https://au.shein.com/1", 5.0, None, "", None, None),          # PENDING
        (2, "https://au.shein.com/2", None, None, "", "2026-01-01", "Done"),  # skip
        (3, "https://au.shein.com/3", None, None, "", None, "Failed"),     # skip
        (4, "https://au.shein.com/4", None, None, "", "2026-01-01", None), # skip
        (5, None, None, None, "", None, None),                             # skip (no URL)
        (6, "https://au.shein.com/6", 20.0, 3.5, "Color:Red", None, None), # PENDING
    ])
    pending = _read_pending_rows(ws)
    seqs = [p["seq"] for p in pending]
    assert seqs == [1, 6], seqs
    p6 = pending[1]
    assert p6["price"] == 20.0
    assert p6["shipping"] == 3.5
    assert p6["variant_filter"] == "Color:Red"


def test_pending_skips_non_template_sheet():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "some other header"
    ws.cell(1, 2).value = "not-链接"
    ws.cell(2, 2).value = "https://au.shein.com/1"
    assert _read_pending_rows(ws) == []


def test_pending_skips_row_with_non_numeric_seq_but_continues():
    wb = Workbook()
    ws = _make_template_ws(wb, [
        ("abc", "https://au.shein.com/1", None, None, "", None, None),  # skip
        (2, "https://au.shein.com/2", None, None, "", None, None),        # PENDING
    ])
    pending = _read_pending_rows(ws)
    assert [p["seq"] for p in pending] == [2]


def test_pending_skips_row_with_non_numeric_price_but_continues():
    wb = Workbook()
    ws = _make_template_ws(wb, [
        (1, "https://au.shein.com/1", "not-a-price", None, "", None, None),  # skip
        (2, "https://au.shein.com/2", 5.5, None, "", None, None),               # PENDING
    ])
    pending = _read_pending_rows(ws)
    assert [p["seq"] for p in pending] == [2]


# ── _backup_workbook + _prune_backups ───────────────────────────────────────

def test_backup_copies_input_into_backup_dir_with_timestamp():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        src = td / "input.xlsx"
        wb = Workbook(); wb.active["A1"] = "marker"; wb.save(src)
        backups = td / "_backups"

        dst = _backup_workbook(src, backups)
        assert dst.parent == backups
        assert dst.name.startswith("input-") and dst.suffix == ".xlsx"
        # Content preserved
        assert load_workbook(dst).active["A1"].value == "marker"
        # Original untouched
        assert load_workbook(src).active["A1"].value == "marker"


def test_prune_keeps_newest_n_and_drops_older():
    with tempfile.TemporaryDirectory() as td:
        backups = Path(td) / "_backups"
        backups.mkdir()
        # 5 backups with staggered mtimes; keep 3
        paths = []
        for i in range(5):
            p = backups / f"input-2026-01-01_00000{i}.xlsx"
            p.write_bytes(b"x")
            # Force mtime to be i seconds apart (increasing)
            t = time.time() - (5 - i) * 60
            os.utime(p, (t, t))
            paths.append(p)
        removed = _prune_backups(backups, "input", keep=3)
        assert removed == 2
        remaining = sorted(p.name for p in backups.glob("input-*.xlsx"))
        # Newest 3 stay: 002, 003, 004
        assert remaining == [
            "input-2026-01-01_000002.xlsx",
            "input-2026-01-01_000003.xlsx",
            "input-2026-01-01_000004.xlsx",
        ], remaining


def test_prune_excludes_legacy_backups():
    """`-legacy-*` files are one-shot migration artefacts; they should stay
    forever until the operator clears them by hand."""
    with tempfile.TemporaryDirectory() as td:
        backups = Path(td) / "_backups"
        backups.mkdir()
        (backups / "input-2026-01-01_000000.xlsx").write_bytes(b"x")
        (backups / "input-2026-01-01_000001.xlsx").write_bytes(b"x")
        legacy = backups / "input-legacy-2026-01-01.xlsx"
        legacy.write_bytes(b"legacy")
        removed = _prune_backups(backups, "input", keep=1)
        # Only one non-legacy pruned; legacy untouched.
        assert removed == 1
        assert legacy.exists()


def test_prune_keep_zero_or_none_is_no_op():
    with tempfile.TemporaryDirectory() as td:
        backups = Path(td) / "_backups"
        backups.mkdir()
        (backups / "input-2026-01-01_000000.xlsx").write_bytes(b"x")
        assert _prune_backups(backups, "input", keep=0) == 0
        assert _prune_backups(backups, "input", keep=None) == 0


# ── _archive_legacy_enriched ────────────────────────────────────────────────

def test_archive_legacy_moves_file_with_dated_suffix():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        old = td / "澳洲希音链接 (输出) - ZR.xlsx"
        old.write_bytes(b"old enriched")
        backups = td / "_backups"

        dst = _archive_legacy_enriched(old, backups)
        assert not old.exists(), "old enriched should have been moved"
        assert dst.parent == backups
        assert "-legacy-" in dst.name
        assert dst.name.endswith(".xlsx")
        assert dst.read_bytes() == b"old enriched"


def test_archive_legacy_avoids_same_day_collision():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        backups = td / "_backups"
        backups.mkdir()
        ts = datetime.now().strftime("%Y-%m-%d")
        # Pre-existing collision on today's date
        (backups / f"file-legacy-{ts}.xlsx").write_bytes(b"prev")
        old = td / "file.xlsx"
        old.write_bytes(b"new content")

        dst = _archive_legacy_enriched(old, backups)
        assert dst.name.endswith(f"-legacy-{ts}-2.xlsx"), dst.name
        assert dst.read_bytes() == b"new content"
        # First collision preserved
        assert (backups / f"file-legacy-{ts}.xlsx").read_bytes() == b"prev"


if __name__ == "__main__":
    test_pending_selects_only_rows_with_url_and_empty_f_g()
    test_pending_skips_non_template_sheet()
    test_pending_skips_row_with_non_numeric_seq_but_continues()
    test_pending_skips_row_with_non_numeric_price_but_continues()
    test_backup_copies_input_into_backup_dir_with_timestamp()
    test_prune_keeps_newest_n_and_drops_older()
    test_prune_excludes_legacy_backups()
    test_prune_keep_zero_or_none_is_no_op()
    test_archive_legacy_moves_file_with_dated_suffix()
    test_archive_legacy_avoids_same_day_collision()
    print("ALL PASS")
