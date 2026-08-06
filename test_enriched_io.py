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


def test_self_heals_partial_and_typo_headers():
    """User provides row 1 with typos, missing cols, or extras — script
    silently rewrites row 1 to the canonical EXPECTED_HEADERS. Data rows
    below (if any) are preserved. Root-cause fix for the 'output empty'
    bug employees hit when their headers didn't match exactly."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "ZR1"
        # Wrong headers: only 3 cols filled, one typo'd
        _write_headers(ws0, ["编号", "链接", "OOPS"])
        # Data row 2 — should be preserved untouched
        ws0.cell(2, 1).value = 999
        ws0.cell(2, 2).value = "http://existing"
        wb0.save(p)

        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        # Row 1 fully seeded to canonical
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            assert ws.cell(1, ci).value == expected, f"col {ci} not healed"
        # Row 2 (existing data) preserved
        assert ws.cell(2, 1).value == 999
        assert ws.cell(2, 2).value == "http://existing"


def test_self_heals_completely_empty_row_1():
    """A brand-new sheet (no headers at all) gets seeded — this used to
    raise ValueError on col 1 (None != '编号')."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "ZR1"
        # No headers, just the empty sheet
        wb0.save(p)

        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            assert ws.cell(1, ci).value == expected, f"col {ci} not seeded"


def test_default_workbook_sheet_removed_when_creating_new_file():
    # openpyxl's Workbook() ships with a default "Sheet"; when we create the
    # enriched file for the first time, that phantom sheet should be gone.
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        wb.save(p)
        wb2 = load_workbook(p)
        assert wb2.sheetnames == ["ZR1"], wb2.sheetnames


def test_append_row_writes_at_max_row_plus_one():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        # Import the appender we're about to build.
        from run_excel import _append_enriched_row

        master_row = {
            "row": 5, "seq": 42, "url": "http://u", "price": 10.0,
            "shipping": 7.95, "variant_filter": "Black",
        }
        # Simulate a successful scrape result.
        result = {
            "date": "2026-08-04",
            "status": "Done",
            "web_price": 9.99,
            "shein_title": "Shein Title",
            "ebay_title": "eBay Title",
            "ebay_price": 25.93,
            "stock": "M: 20",
        }
        _append_enriched_row(ws, master_row, result, picture_path=None)
        # First data row = row 2.
        assert ws.cell(2, 1).value == 42       # A 编号
        assert ws.cell(2, 2).value == "http://u"  # B 链接
        assert ws.cell(2, 3).value == 10.0     # C 原价
        assert ws.cell(2, 4).value == 7.95     # D 运费
        assert ws.cell(2, 5).value == "Black"  # E 变体
        assert ws.cell(2, 6).value == "2026-08-04"  # F 日期
        assert ws.cell(2, 7).value == "Done"    # G 状态
        # H 图片 — no picture_path → cell empty (image anchored elsewhere)
        assert ws.cell(2, 8).value is None
        assert ws.cell(2, 9).value == 9.99      # I 希音价格
        assert ws.cell(2, 10).value == "Shein Title"  # J
        assert ws.cell(2, 11).value == "eBay Title"   # K
        assert ws.cell(2, 12).value == 25.93    # L eBay 价格
        assert ws.cell(2, 13).value == "M: 20"  # M 库存

        # Second append lands at row 3 (max_row+1), not overwrite row 2.
        _append_enriched_row(ws, master_row, result, picture_path=None)
        assert ws.cell(2, 1).value == 42  # still there
        assert ws.cell(3, 1).value == 42  # second appended


def test_append_row_on_failed_scrape_writes_only_date_and_status():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        from run_excel import _append_enriched_row

        master_row = {
            "row": 3, "seq": 7, "url": "http://ux", "price": 5.0,
            "shipping": 0.0, "variant_filter": "",
        }
        result = {"date": "2026-08-04", "status": "Failed"}
        _append_enriched_row(ws, master_row, result, picture_path=None)
        assert ws.cell(2, 1).value == 7
        assert ws.cell(2, 6).value == "2026-08-04"
        assert ws.cell(2, 7).value == "Failed"
        # Result cols (H-M) should be None on failure.
        for c in range(8, 14):
            assert ws.cell(2, c).value is None


def test_compact_pulls_data_up_below_header():
    """User's file has header + 990 empty rows + 10 data rows at 993-1002.
    After compaction, data lands at rows 2-11 (glued to header) and
    ws.max_row drops to 11.
    """
    from run_excel import _compact_data_rows

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        # Simulate user's messy file: 10 data rows starting at row 993.
        for i, r in enumerate(range(993, 1003), start=1):
            ws.cell(r, 1).value = i           # seq
            ws.cell(r, 6).value = f"2026-08-0{i%9+1}"
            ws.cell(r, 7).value = "Done"
        assert ws.max_row >= 1002

        removed = _compact_data_rows(ws, key_col=1)
        # Removed 991 rows total: 991 empty rows between row 1 and row 1002.
        assert removed == 991, removed
        # Data now at rows 2-11.
        assert ws.max_row == 11
        for i, r in enumerate(range(2, 12), start=1):
            assert ws.cell(r, 1).value == i, (
                f"row {r} col A expected {i}, got {ws.cell(r,1).value!r}")


def test_compact_removes_trailing_empty_formatting_only():
    """When data starts right below header but trailing formatting extends
    max_row past it, only the trailing empty rows are removed."""
    from openpyxl.styles import Border, Side
    from run_excel import _compact_data_rows

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        ws.cell(2, 1).value = 1
        ws.cell(3, 1).value = 2
        # Formatting at row 500 → bumps max_row to 500
        ws.cell(500, 5).border = Border(top=Side(style="thin"))
        assert ws.max_row >= 500

        removed = _compact_data_rows(ws, key_col=1)
        assert removed >= 495, removed
        # Data still at rows 2-3
        assert ws.cell(2, 1).value == 1
        assert ws.cell(3, 1).value == 2


def test_ensure_enriched_sheet_auto_compacts_on_open():
    """_ensure_enriched_sheet compacts as part of opening — no separate call
    needed. Verifies the end-to-end UX employees see."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        # Create a mock-messy file: headers + empty rows + late data.
        wb0 = Workbook()
        ws0 = wb0.active
        ws0.title = "ZR1"
        _write_headers(ws0, EXPECTED_HEADERS)
        ws0.cell(500, 1).value = 42     # data buried at row 500
        ws0.cell(500, 7).value = "Done"
        wb0.save(p)

        # Open via our helper — should auto-compact.
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        assert ws.cell(2, 1).value == 42, (
            f"data not pulled up; row 2 col A = {ws.cell(2,1).value!r}")


def test_append_ignores_phantom_empty_rows_from_formatting():
    """Reproduce the 'output empty' bug: operator's fresh xlsx has phantom
    empty rows extending far past the header (Excel formatting artifact).
    ws.max_row reports a high number, but no real data. Append must land
    at row 2, not row 1001.
    """
    from openpyxl.styles import Border, Side
    from run_excel import _append_enriched_row, _next_data_row

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        # Simulate the artifact: apply formatting to a distant empty row
        # (this bumps openpyxl's ws.max_row without adding any data).
        border = Border(left=Side(style="thin"))
        for r in range(50, 101):
            for c in range(1, 22):
                ws.cell(r, c).border = border
        # openpyxl.max_row should now report >= 100 despite zero data rows
        assert ws.max_row >= 100, f"formatting didn't bump max_row: got {ws.max_row}"
        # Our helper should still know: next data row is 2 (no data yet).
        assert _next_data_row(ws, key_col=1) == 2

        # Append: should land at row 2, NOT row max_row+1.
        master_row = {"row": 5, "seq": 99, "url": "http://u",
                      "price": None, "shipping": None, "variant_filter": ""}
        _append_enriched_row(ws, master_row, {"date": "d", "status": "Done"})
        assert ws.cell(2, 1).value == 99, (
            f"append landed at wrong row; row 2 col A = {ws.cell(2,1).value!r}, "
            f"max_row = {ws.max_row}")


def test_append_after_existing_data_ignores_further_phantom_rows():
    """If real data exists at row 5 but phantom formatting extends to row
    500, the next append lands at row 6 — right after the last real row."""
    from openpyxl.styles import Border, Side
    from run_excel import _append_enriched_row, _next_data_row

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "enriched.xlsx"
        wb, ws = _ensure_enriched_sheet(p, "ZR1")
        # Real data at row 5
        ws.cell(5, 1).value = 42
        # Formatting at row 500 → bumps max_row
        ws.cell(500, 1).border = Border(top=Side(style="thin"))
        assert ws.max_row >= 500
        assert _next_data_row(ws, key_col=1) == 6

        master_row = {"row": 8, "seq": 77, "url": "http://u",
                      "price": None, "shipping": None, "variant_filter": ""}
        _append_enriched_row(ws, master_row, {"date": "d", "status": "Done"})
        assert ws.cell(6, 1).value == 77


if __name__ == "__main__":
    test_creates_file_and_sheet_when_absent()
    test_creates_sheet_when_file_exists_but_sheet_missing()
    test_reuses_existing_sheet_with_matching_headers()
    test_self_heals_partial_and_typo_headers()
    test_self_heals_completely_empty_row_1()
    test_default_workbook_sheet_removed_when_creating_new_file()
    test_append_row_writes_at_max_row_plus_one()
    test_append_row_on_failed_scrape_writes_only_date_and_status()
    test_compact_pulls_data_up_below_header()
    test_compact_removes_trailing_empty_formatting_only()
    test_ensure_enriched_sheet_auto_compacts_on_open()
    test_append_ignores_phantom_empty_rows_from_formatting()
    test_append_after_existing_data_ignores_further_phantom_rows()
    print("ALL PASS")
