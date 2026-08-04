"""Tests for _ensure_enriched_sheet — loads or creates the enriched workbook,
loads or creates the store sheet with the 21-col header, and validates
existing headers to prevent silent schema drift."""
import tempfile
from pathlib import Path
from openpyxl import Workbook, load_workbook

from run_excel import _ensure_enriched_sheet, EXPECTED_HEADERS


def _write_headers(ws, headers):
    for ci, h in enumerate(headers, 1):
        ws.cell(1, ci).value = h


def test_creates_file_and_sheet_when_absent():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        # File does not exist yet.
        assert not p.exists()
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        assert ws.title == "ZR1"
        # Header row seeded.
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            assert ws.cell(1, ci).value == expected, f"col {ci}: expected {expected!r}"
        wb.save(p)
        assert p.exists()


def test_creates_sheet_when_file_exists_but_sheet_missing():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        # Seed file with a different sheet (headers valid on that one).
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "OTHER"
        _write_headers(ws0, EXPECTED_HEADERS)
        wb0.save(p)

        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        assert "ZR1" in wb.sheetnames
        assert "OTHER" in wb.sheetnames  # other sheet preserved
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            assert ws.cell(1, ci).value == expected


def test_reuses_existing_sheet_with_matching_headers():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "ZR1"
        _write_headers(ws0, EXPECTED_HEADERS)
        # Seed one data row.
        ws0.cell(2, 1).value = 1
        ws0.cell(2, 2).value = "http://u1"
        wb0.save(p)

        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        assert ws.cell(2, 1).value == 1  # existing data untouched
        assert ws.cell(2, 2).value == "http://u1"


def test_hard_error_on_header_mismatch():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "ZR1"
        # Wrong headers (missing 图片, extras)
        _write_headers(ws0, ["编号", "链接", "OOPS"])
        wb0.save(p)

        raised = False
        try:
            _ensure_enriched_sheet(p, "ZR1")
        except ValueError as e:
            raised = True
            assert "header" in str(e).lower(), str(e)
        assert raised, "expected ValueError on header mismatch"


def test_default_workbook_sheet_removed_when_creating_new_file():
    # openpyxl's Workbook() ships with a default "Sheet"; when we create the
    # enriched file for the first time, that phantom sheet should be gone.
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        wb.save(p)
        wb2 = load_workbook(p)
        assert wb2.sheetnames == ["ZR1"], wb2.sheetnames


if __name__ == "__main__":
    test_creates_file_and_sheet_when_absent()
    test_creates_sheet_when_file_exists_but_sheet_missing()
    test_reuses_existing_sheet_with_matching_headers()
    test_hard_error_on_header_mismatch()
    test_default_workbook_sheet_removed_when_creating_new_file()
    print("ALL PASS")
