# Master/Enriched File Split Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split the current single-file template into two files — a read-only master (6 cols, user-owned, F=Y trigger) and an append-only enriched output (21 cols matching current template, script-owned). Corruption of enriched loses only append history since last backup; master is untouchable by any script.

**Architecture:**
- `run_excel.py` reads master (never writes it), scrapes per F=Y trigger, and APPENDS one row per (product × run) to the enriched file. Enriched grows over time — the same seq processed 5 times shows 5 rows with different `F 日期`.
- `ebay_price_check.py` reads enriched (not master), updates N-R in-place for rows with `G=Done AND N=empty AND K non-empty`.
- Two `.env` keys drive filenames: `SHEIN_INPUT_FILENAME` (master, required) and `SHEIN_OUTPUT_FILENAME` (enriched, required). `setup_wizard.py` collects both.
- Enriched file/sheet auto-created on first run (headers seeded). Header mismatch on an existing file/sheet is a hard error to prevent schema drift.

**Tech Stack:** Python 3.10+, openpyxl, existing CDP infrastructure. Plain-`assert` tests (project convention — see `test_variant_merge.py`).

**Working branch:** `feat/master-enriched-split` (already checked out at commit `560bd42`).

**Reference spec:** `docs/superpowers/specs/2026-08-04-master-enriched-split-design.md`.

**Version bump:** 0.2.1 → 0.3.0 (minor — new required config key, breaks old single-file workflow).

---

## File Structure

**New files:**
- `test_master_reader.py` — tests for `_read_master_pending_rows()` (new master schema: 6 cols, F=Y trigger)
- `test_enriched_io.py` — tests for enriched sheet lifecycle + append helper

**Modified files:**
- `config.py` — add `OUTPUT_FILENAME` + `_require_output_filename()` helper
- `run_excel.py`
  - Add `MASTER_COL_TRIGGER=6` constant + `MASTER_EXPECTED_HEADERS`
  - Rewrite `_read_pending_rows()` → `_read_master_pending_rows()` (6-col master, F=Y filter)
  - Add `_ensure_enriched_sheet(wb, sheet_name) -> ws`
  - Add `_append_enriched_row(ws, master_row_dict, result_dict, picture_path)`
  - Rewrite `process_excel()` to read master + append to enriched (never write master)
  - Update `main()` to require + pass both filenames
  - Module docstring updated for new architecture
- `ebay_price_check.py`
  - Change file target from `INPUT_FILENAME` → `OUTPUT_FILENAME` (enriched)
  - Reader logic unchanged (enriched has the same 21-col schema as the old single file)
- `setup_wizard.py`
  - Rename "指定输入文件名（可选）" → "输入表文件名（必填）"
  - Add "输出表文件名（必填）" field
- `version.py` — `0.2.1` → `0.3.0`
- `installer.iss` — `MyAppVersion "0.2.1"` → `"0.3.0"`
- `.env` (per-machine, gitignored) — add `SHEIN_OUTPUT_FILENAME=澳洲希音链接 (输出) - ZR.xlsx`

**Deleted files:** none.

---

## Task 1: Add OUTPUT_FILENAME config + fail-loud helper

**Files:**
- Modify: `config.py`
- Modify: `.env` (local machine only, gitignored)

- [ ] **Step 1: Add OUTPUT_FILENAME and helper to config.py**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\config.py`. Find the block at the bottom (after `OUTPUT_ROOT_2ND`):

```python
OUTPUT_ROOT_2ND = Path(os.environ.get(
    "SHEIN_OUTPUT_DIR",
    _AU_BASE,  # 默认与 SUBMITTED_DIR 同级；店铺名自动作为下一级子文件夹
))
```

Append after it:

```python
# ── Enriched (输出) 文件名 ──────────────────────────────────────────────────
# 主表(输入)与富表(输出)分开后新增。SHEIN_OUTPUT_FILENAME 必填，无默认。
# 具体检查在使用点 (require_output_filename) 抛出，避免 config 载入即失败。
OUTPUT_FILENAME = os.environ.get("SHEIN_OUTPUT_FILENAME", "").strip()


def require_output_filename() -> str:
    """Return OUTPUT_FILENAME or exit with a clear message.

    Called by run_excel.py / ebay_price_check.py at startup — keeps
    `import config` side-effect-free for test contexts."""
    if not OUTPUT_FILENAME:
        import sys
        sys.stderr.write(
            "错误: SHEIN_OUTPUT_FILENAME 未在 .env / config.env 中设置。\n"
            "请重新运行 setup_wizard 并填写'输出表文件名'，或手动编辑\n"
            "%APPDATA%\\shein-extract-au\\config.env 加一行\n"
            "  SHEIN_OUTPUT_FILENAME=<你的输出表文件名.xlsx>\n"
        )
        sys.exit(1)
    return OUTPUT_FILENAME
```

- [ ] **Step 2: Update local .env**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\.env`. It should already have `SHEIN_INPUT_FILENAME=澳洲希音链接汇总 - ZR.xlsx` from a prior task. Change that line to the new master filename AND add the output line:

Replace:
```
SHEIN_INPUT_FILENAME=澳洲希音链接汇总 - ZR.xlsx
```

With (two lines):
```
SHEIN_INPUT_FILENAME=澳洲希音链接 (输入) - ZR.xlsx
SHEIN_OUTPUT_FILENAME=澳洲希音链接 (输出) - ZR.xlsx
```

- [ ] **Step 3: Verify config resolves**

Run:
```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from config import INPUT_FILENAME, OUTPUT_FILENAME, require_output_filename, SUBMITTED_DIR
print(f'INPUT_FILENAME  -> {INPUT_FILENAME}')
print(f'OUTPUT_FILENAME -> {OUTPUT_FILENAME}')
print(f'input path      -> {(SUBMITTED_DIR / INPUT_FILENAME).exists()=}')
print(f'output path     -> {(SUBMITTED_DIR / OUTPUT_FILENAME).exists()=}')
print(f'require_output_filename() -> {require_output_filename()}')
"
```

Expected: both filenames shown, `require_output_filename()` returns the OUTPUT_FILENAME (no exit).

- [ ] **Step 4: Commit config change (do not commit .env — it's gitignored)**

```bash
cd "C:\Users\ak\Desktop\Claude\shein-extract-au"
git add config.py
git commit -m "feat(config): add SHEIN_OUTPUT_FILENAME + require_output_filename() helper"
```

---

## Task 2: TDD master reader (6-col schema + F=Y trigger)

**Files:**
- Create: `test_master_reader.py`
- Modify: `run_excel.py` — add `MASTER_COL_TRIGGER` constant + `MASTER_EXPECTED_HEADERS` + `_read_master_pending_rows()` function

- [ ] **Step 1: Write failing tests**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_master_reader.py`:

```python
"""Tests for _read_master_pending_rows — reads the 6-col master schema and
returns rows where F (是否要跑) normalizes to 'Y' (case + whitespace tolerant)."""
from openpyxl import Workbook

from run_excel import _read_master_pending_rows


MASTER_HEADERS = ["编号", "链接", "原价", "运费", "变体", "是否要跑"]
NCOLS = len(MASTER_HEADERS)  # 6


def _make_ws(rows: list):
    wb = Workbook()
    ws = wb.active
    ws.title = "ZR1"
    for ci, h in enumerate(MASTER_HEADERS, 1):
        ws.cell(1, ci).value = h
    for ri, r in enumerate(rows, 2):
        padded = list(r) + [None] * (NCOLS - len(r))
        for ci, v in enumerate(padded, 1):
            ws.cell(ri, ci).value = v
    return ws


def test_y_uppercase_triggers():
    ws = _make_ws([
        [1, "http://u1", 10.0, 7.95, "", "Y"],
    ])
    pending = _read_master_pending_rows(ws)
    assert len(pending) == 1
    assert pending[0]["row"] == 2
    assert pending[0]["seq"] == 1
    assert pending[0]["url"] == "http://u1"
    assert pending[0]["price"] == 10.0
    assert pending[0]["shipping"] == 7.95
    assert pending[0]["variant_filter"] == ""


def test_y_lowercase_and_whitespace_triggers():
    ws = _make_ws([
        [1, "http://u1", 10.0, 7.95, "", "y"],
        [2, "http://u2", 10.0, 7.95, "", "  Y  "],
        [3, "http://u3", 10.0, 7.95, "", "yes"],  # only exact Y (any case) triggers
    ])
    pending = _read_master_pending_rows(ws)
    seqs = [p["seq"] for p in pending]
    assert seqs == [1, 2], seqs  # 'yes' does NOT trigger; only 'y' / 'Y'


def test_non_y_is_skipped():
    ws = _make_ws([
        [1, "http://u1", 10.0, 7.95, "", ""],       # empty
        [2, "http://u2", 10.0, 7.95, "", "N"],       # explicit skip
        [3, "http://u3", 10.0, 7.95, "", "完成"],    # arbitrary marker
        [4, "http://u4", 10.0, 7.95, "", None],      # None
    ])
    pending = _read_master_pending_rows(ws)
    assert pending == []


def test_variant_filter_and_optional_price_shipping():
    ws = _make_ws([
        [1, "http://u1", None, None, "Black, Red / M, L", "Y"],
    ])
    pending = _read_master_pending_rows(ws)
    assert len(pending) == 1
    p = pending[0]
    assert p["price"] is None
    assert p["shipping"] is None
    assert p["variant_filter"] == "Black, Red / M, L"


def test_non_master_sheet_returns_empty():
    # A sheet whose B1 is not '链接' isn't the master schema.
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "Seq"
    ws.cell(1, 2).value = "Website"
    ws.cell(2, 1).value = 1
    ws.cell(2, 6).value = "Y"
    assert _read_master_pending_rows(ws) == []


def test_non_numeric_seq_skipped():
    ws = _make_ws([
        ["oops", "http://u1", 10.0, 7.95, "", "Y"],
        [2, "http://u2", 10.0, 7.95, "", "Y"],
    ])
    pending = _read_master_pending_rows(ws)
    assert [p["seq"] for p in pending] == [2]


def test_missing_url_skipped():
    ws = _make_ws([
        [1, None, 10.0, 7.95, "", "Y"],
        [2, "http://u2", 10.0, 7.95, "", "Y"],
    ])
    pending = _read_master_pending_rows(ws)
    assert [p["seq"] for p in pending] == [2]


if __name__ == "__main__":
    test_y_uppercase_triggers()
    test_y_lowercase_and_whitespace_triggers()
    test_non_y_is_skipped()
    test_variant_filter_and_optional_price_shipping()
    test_non_master_sheet_returns_empty()
    test_non_numeric_seq_skipped()
    test_missing_url_skipped()
    print("ALL PASS")
```

- [ ] **Step 2: Confirm RED**

```bash
python test_master_reader.py
```
Expected: `ImportError: cannot import name '_read_master_pending_rows' from 'run_excel'`.

- [ ] **Step 3: Add constants and function to run_excel.py**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`. Find the block after `EXPECTED_HEADERS`:

```python
EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体",
                    "日期", "状态", "图片", "希音价格", "希音标题",
                    "eBay标题", "eBay价格", "库存"]
```

Insert AFTER it (before `def _sheet_matches_template`):

```python

# ── Master (输入) schema ─────────────────────────────────────────────────────
# 主表只有 6 列，脚本只读。See spec 2026-08-04-master-enriched-split-design.md §1.
MASTER_COL_TRIGGER = 6  # F 是否要跑 (Y/空/其他)

MASTER_EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体", "是否要跑"]


def _read_master_pending_rows(ws) -> list[dict]:
    """Return list of dicts for rows in the master where F (是否要跑) normalizes
    to 'Y' (case-insensitive, whitespace-stripped). Non-master sheets return [].
    Shape: {row, seq, url, price, shipping, variant_filter}. Rows with invalid
    seq or missing URL are logged and skipped."""
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        url = ws.cell(r, COL_URL).value
        trigger = str(ws.cell(r, MASTER_COL_TRIGGER).value or "").strip().upper()
        if trigger != "Y":
            continue
        if not url:
            logger.info("  row %d: skip (F=Y but 链接 empty)", r)
            continue
        try:
            seq_int = int(seq) if seq is not None else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (编号 '%s' not numeric)", r, seq)
            continue
        if seq_int is None:
            logger.info("  row %d: skip (编号 empty)", r)
            continue
        raw_price = ws.cell(r, COL_PRICE).value
        raw_ship = ws.cell(r, COL_SHIPPING).value
        try:
            price = float(raw_price) if raw_price not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (原价 '%s' not numeric)", r, raw_price)
            continue
        try:
            shipping = float(raw_ship) if raw_ship not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (运费 '%s' not numeric)", r, raw_ship)
            continue
        variant_filter = str(ws.cell(r, COL_VARIANT_FILTER).value or "").strip()
        pending.append({
            "row": r,
            "seq": seq_int,
            "url": str(url).strip(),
            "price": price,
            "shipping": shipping,
            "variant_filter": variant_filter,
        })
    return pending
```

Note: `_sheet_matches_template` (already defined above) checks `B1 == "链接"` — same rule works for master since master's B is also 链接.

- [ ] **Step 4: Confirm GREEN**

```bash
python test_master_reader.py
```
Expected: `ALL PASS`.

Also confirm no regression on other tests:
```bash
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
```
Expected: four `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add test_master_reader.py run_excel.py
git commit -m "feat(run_excel): add master-schema reader (_read_master_pending_rows) — F=Y trigger"
```

---

## Task 3: TDD enriched sheet lifecycle helpers

**Files:**
- Create: `test_enriched_io.py`
- Modify: `run_excel.py` — add `_ensure_enriched_sheet()` function

- [ ] **Step 1: Write failing tests**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_enriched_io.py`:

```python
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
```

- [ ] **Step 2: Confirm RED**

```bash
python test_enriched_io.py
```
Expected: `ImportError: cannot import name '_ensure_enriched_sheet' from 'run_excel'`.

- [ ] **Step 3: Add helper to run_excel.py**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`. After `_read_master_pending_rows` (added in Task 2), add:

```python
def _ensure_enriched_sheet(enriched_path, sheet_name: str):
    """Load or create the enriched workbook, then load or create the store
    sheet with the 21-col header row. Returns (wb, ws). Raises ValueError if
    a sheet exists with mismatched headers (prevents silent schema drift)."""
    from pathlib import Path
    from openpyxl import Workbook, load_workbook

    p = Path(enriched_path)
    if p.exists():
        wb = load_workbook(p)
    else:
        wb = Workbook()
        # openpyxl seeds a default 'Sheet' we don't want in the final layout.
        default_name = wb.sheetnames[0]
        if default_name != sheet_name:
            del wb[default_name]

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Validate headers on existing sheet — hard error on drift.
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            actual = ws.cell(1, ci).value
            if actual != expected:
                raise ValueError(
                    f"Enriched sheet '{sheet_name}' header mismatch at col {ci}: "
                    f"expected {expected!r}, got {actual!r}. Fix the file or "
                    f"delete the sheet to have it recreated."
                )
    else:
        ws = wb.create_sheet(sheet_name)
        for ci, header in enumerate(EXPECTED_HEADERS, 1):
            ws.cell(1, ci).value = header

    return wb, ws
```

- [ ] **Step 4: Confirm GREEN**

```bash
python test_enriched_io.py
```
Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add test_enriched_io.py run_excel.py
git commit -m "feat(run_excel): _ensure_enriched_sheet loads/creates enriched wb+sheet with header validation"
```

---

## Task 4: TDD enriched row appender

**Files:**
- Modify: `test_enriched_io.py` — extend with append tests
- Modify: `run_excel.py` — add `_append_enriched_row()` function

- [ ] **Step 1: Extend test_enriched_io.py with append tests**

Add these tests BEFORE the `if __name__` block:

```python
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
```

Add their calls to the `if __name__ == "__main__":` block:

```python
    test_append_row_writes_at_max_row_plus_one()
    test_append_row_on_failed_scrape_writes_only_date_and_status()
```

- [ ] **Step 2: Confirm RED**

```bash
python test_enriched_io.py
```
Expected: `ImportError: cannot import name '_append_enriched_row' from 'run_excel'`.

- [ ] **Step 3: Add _append_enriched_row to run_excel.py**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`. After `_ensure_enriched_sheet` (added in Task 3), add:

```python
def _append_enriched_row(ws, master_row: dict, result: dict,
                          picture_path=None) -> int:
    """Append one row to the enriched sheet at row = ws.max_row + 1.

    - master_row must have: seq, url, price, shipping, variant_filter
      (copied verbatim to A-E — a snapshot of the master row's values at
      scrape time).
    - result may have: date, status, web_price, shein_title, ebay_title,
      ebay_price, stock. Unspecified keys skip that cell.
    - picture_path (optional) — file path to embed in H; row height auto-grown.

    Returns the row index that was written."""
    row = ws.max_row + 1
    # If the sheet only has the header row, max_row is 1 → row = 2 (correct).
    # If it has header + N data rows, row = N + 2 (correct).

    # A-E: master snapshot.
    ws.cell(row, COL_SEQ).value = master_row.get("seq")
    ws.cell(row, COL_URL).value = master_row.get("url")
    if master_row.get("price") is not None:
        ws.cell(row, COL_PRICE).value = master_row["price"]
    if master_row.get("shipping") is not None:
        ws.cell(row, COL_SHIPPING).value = master_row["shipping"]
    if master_row.get("variant_filter"):
        ws.cell(row, COL_VARIANT_FILTER).value = master_row["variant_filter"]

    # F-M: scrape result (via existing writer, which already handles
    # None-means-skip semantics and image embedding).
    _write_result_row(
        ws, row=row,
        date=result.get("date", ""),
        status=result.get("status", ""),
        picture_path=picture_path,
        web_price=result.get("web_price"),
        shein_title=result.get("shein_title"),
        ebay_title=result.get("ebay_title"),
        ebay_price=result.get("ebay_price"),
        stock=result.get("stock"),
    )
    return row
```

Note: `COL_SEQ`, `COL_URL`, `COL_PRICE`, `COL_SHIPPING`, `COL_VARIANT_FILTER`, and `_write_result_row` are already defined earlier in `run_excel.py` from prior refactor work.

- [ ] **Step 4: Confirm GREEN**

```bash
python test_enriched_io.py
```
Expected: `ALL PASS` (5 tests: 3 lifecycle + 2 append).

Also re-run master reader tests:
```bash
python test_master_reader.py
```
Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add test_enriched_io.py run_excel.py
git commit -m "feat(run_excel): _append_enriched_row writes master A-E snapshot + scrape result to next row"
```

---

## Task 5: Rewire `process_excel()` for master → enriched flow

**Files:**
- Modify: `run_excel.py::process_excel` (major rewrite)
- Modify: `run_excel.py::main` (pass both filenames)

- [ ] **Step 1: Rewrite process_excel signature and body**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`. Find the current `def process_excel(xlsx_path: Path) -> None:` function. Replace its entire body with:

```python
def process_excel(master_path: Path, enriched_path: Path) -> None:
    """Read the master (input) workbook, scrape every row where F='Y', and
    append one row per (product × run) to the enriched (output) workbook.
    The master is opened read-only; only the enriched is saved. See spec
    2026-08-04-master-enriched-split-design.md."""
    logger.info("Master:   %s", master_path.name)
    logger.info("Enriched: %s", enriched_path.name)

    if not master_path.exists():
        logger.error("Master file not found: %s", master_path)
        return

    master_wb = load_workbook(master_path, read_only=False)  # read-only intent — no save() called

    for ws_name in master_wb.sheetnames:
        master_ws = master_wb[ws_name]
        store = ws_name.strip()
        pending = _read_master_pending_rows(master_ws)
        if not pending:
            logger.info("  Sheet '%s': no F=Y rows (or not master schema)", store)
            continue

        logger.info("  Sheet '%s': %d row(s) with F=Y: seq %s",
                    store, len(pending), [p["seq"] for p in pending])

        store_dir = OUTPUT_ROOT / store
        store_dir.mkdir(parents=True, exist_ok=True)

        today = datetime.now().strftime("%Y-%m-%d")
        seqs = [p["seq"] for p in pending]
        seq_min, seq_max = min(seqs), max(seqs)
        batch_xlsx = f"{store}-{seq_min}-{seq_max}-{today.replace('-', '')}.xlsx"

        urls = [p["url"] for p in pending]
        prices = [p["price"] for p in pending]
        shippings = [p["shipping"] for p in pending]
        variant_filters = [p["variant_filter"] for p in pending]

        old_cwd = Path.cwd()
        results = None
        try:
            for _retry in range(5):
                try:
                    os.chdir(store_dir)
                    break
                except PermissionError:
                    time.sleep(2)
            else:
                os.chdir(store_dir)
            logger.info("  Scraping %d URLs → %s/%s", len(urls), store, batch_xlsx)
            results = scrape_shein(
                urls,
                output=batch_xlsx,
                seq_list=seqs,
                price_list=prices,
                shipping_list=shippings,
                variant_filter_list=variant_filters,
            )
        except RateLimitError:
            logger.warning("  [限流] Rate limited during '%s'", store)
        except Exception as e:
            logger.exception("  Error processing '%s': %s", store, e)
            try:
                tb = traceback.format_exc()
                (DEBUG_LOG_DIR / "last_traceback.txt").write_text(tb, encoding="utf-8")
            except OSError:
                pass
        finally:
            os.chdir(old_cwd)

        # Open (or create) enriched sheet — one save per sheet at the end.
        try:
            enriched_wb, enriched_ws = _ensure_enriched_sheet(enriched_path, store)
        except ValueError as e:
            logger.error("  Cannot open enriched sheet '%s': %s", store, e)
            continue

        # Append one row per pending master row.
        for p in pending:
            seq = p["seq"]
            seq_folder = store_dir / str(seq)
            has_files = seq_folder.is_dir() and any(seq_folder.iterdir())

            rec = None
            if results:
                rec = next((r for r in results if r.get("seq_num") == seq), None)

            is_bad_data = (rec and rec.get("status") == "OK"
                           and (not rec.get("sku")
                                or "[goods_name]" in (rec.get("title") or "")))

            if has_files and not is_bad_data and rec:
                result_dict = {
                    "date": today,
                    "status": "Done",
                    "web_price": rec.get("web_price_display"),
                    "shein_title": rec.get("original_title") or rec.get("title"),
                    "ebay_title": rec.get("ebay_title"),
                    "ebay_price": rec.get("ebay_price"),
                    "stock": rec.get("stock_summary"),
                }
                picture = rec.get("first_image_path") or None
                _append_enriched_row(enriched_ws, p, result_dict, picture_path=picture)
                logger.info("    seq %d → Done (appended row %d)", seq, enriched_ws.max_row)
            elif rec and rec.get("status") == "DELISTED":
                _append_enriched_row(enriched_ws, p,
                                     {"date": today, "status": "Delisted"})
                logger.info("    seq %d → Delisted", seq)
            else:
                detail = rec.get("status", "") if rec else ""
                if is_bad_data:
                    detail = "no data loaded"
                _append_enriched_row(enriched_ws, p,
                                     {"date": today, "status": "Failed"})
                logger.info("    seq %d → Failed %s", seq,
                            f"({detail})" if detail else "")

        safe_save(enriched_wb, enriched_path)
        logger.info("  Saved enriched progress to %s", enriched_path.name)

    master_wb.close()
    logger.info("Done.")
```

- [ ] **Step 2: Rewrite main() to require + pass both filenames**

Find `def main():` at the bottom of `run_excel.py`. Replace its entire body with:

```python
def main():
    parser = argparse.ArgumentParser(
        description="Master → Enriched Shein scrape pipeline (澳洲站)")
    parser.add_argument("master_file", nargs="?", default=None,
                        help="Master (input) .xlsx path. Default: "
                             "SUBMITTED_DIR/SHEIN_INPUT_FILENAME.")
    parser.add_argument("--enriched", default=None,
                        help="Enriched (output) .xlsx path. Default: "
                             "SUBMITTED_DIR/SHEIN_OUTPUT_FILENAME.")
    args = parser.parse_args()

    setup_logging()

    from config import require_output_filename
    output_name = require_output_filename()

    if args.master_file:
        master_path = Path(args.master_file)
    elif INPUT_FILENAME:
        master_path = SUBMITTED_DIR / INPUT_FILENAME
    else:
        logger.error("No master file given and SHEIN_INPUT_FILENAME not set in .env")
        return

    if args.enriched:
        enriched_path = Path(args.enriched)
    else:
        enriched_path = SUBMITTED_DIR / output_name

    logger.info("=" * 60)
    try:
        process_excel(master_path, enriched_path)
    except Exception as e:
        logger.exception("Fatal error: %s", e)
    logger.info("All done.")
```

- [ ] **Step 3: Verify imports + tests still pass**

```bash
python -c "import run_excel; print('ok')"
python test_master_reader.py
python test_enriched_io.py
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
```

Expected: `ok` and six `ALL PASS`.

Note: `test_template_io.py` may fail on `_read_pending_rows` if we removed/renamed it. Check — if it still exists, tests pass. If we removed it, we need to delete `test_template_io.py`. **Decision: keep `_read_pending_rows` as an alias for backwards-compat during transition. Add this line right after `def _read_master_pending_rows` in run_excel.py:**

Actually simpler: delete `test_template_io.py` — the old reader semantics are gone. Add this step:

- [ ] **Step 4: Delete obsolete test file**

```bash
git rm test_template_io.py
```

This test covered the OLD single-file schema (F=日期, G=状态 as write targets). Under the new architecture, master's F=是否要跑 and enriched's F=日期 are separate concerns, each covered by `test_master_reader.py` and `test_enriched_io.py` respectively.

Also the old `_read_pending_rows` and `_write_result_row` — `_write_result_row` is still needed (called by `_append_enriched_row`), so keep it. `_read_pending_rows` is superseded by `_read_master_pending_rows` — delete it from run_excel.py.

Find and delete the function `def _read_pending_rows(ws) -> list[dict]:` in run_excel.py (currently around line 60-105). It's replaced by `_read_master_pending_rows`.

- [ ] **Step 5: Re-run tests**

```bash
python test_master_reader.py
python test_enriched_io.py
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
```

Expected: five `ALL PASS`.

- [ ] **Step 6: Commit**

```bash
git add run_excel.py
git rm test_template_io.py
git commit -m "refactor(run_excel): process_excel reads master, appends to enriched; drop write-to-master path

- process_excel(master_path, enriched_path) — new signature.
- Uses _read_master_pending_rows (F=Y trigger) + _append_enriched_row.
- Removed obsolete _read_pending_rows (superseded).
- Removed obsolete test_template_io.py (schema fully replaced by master + enriched tests).
- main() requires SHEIN_OUTPUT_FILENAME via require_output_filename()."
```

---

## Task 6: Rewire `ebay_price_check.py` to read/write the enriched file

**Files:**
- Modify: `ebay_price_check.py`

The reader logic (`_read_ebay_pending_rows`) already targets the 21-col schema — which is the enriched schema. Only the FILE PATH needs to change: point at OUTPUT_FILENAME (enriched), not INPUT_FILENAME (master).

- [ ] **Step 1: Update the config resolution in main()**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\ebay_price_check.py`. Find the `main()` function's file-resolution block:

```python
    if args.file:
        xlsx_path = Path(args.file)
    elif INPUT_FILENAME:
        xlsx_path = SUBMITTED_DIR / INPUT_FILENAME
    else:
        logger.error("No file given and SHEIN_INPUT_FILENAME not set in .env")
        sys.exit(1)
```

Change to:

```python
    from config import require_output_filename
    output_name = require_output_filename()

    if args.file:
        xlsx_path = Path(args.file)
    else:
        xlsx_path = SUBMITTED_DIR / output_name
```

Also change the `--help` string for the `file` argument (near the top of `main`):

```python
    parser.add_argument("file", nargs="?", default=None,
                        help="Path to enriched .xlsx (default: SHEIN_OUTPUT_FILENAME under SUBMITTED_DIR)")
```

- [ ] **Step 2: Update the imports block**

Find in ebay_price_check.py:

```python
from config import SUBMITTED_DIR, INPUT_FILENAME
```

Change to (drop INPUT_FILENAME since we no longer use it):

```python
from config import SUBMITTED_DIR
```

- [ ] **Step 3: Update the module docstring**

Near the top of the file, find the docstring that mentions SHEIN_INPUT_FILENAME. Change:

```
Usage:
    python ebay_price_check.py                  # picks SUBMITTED_DIR/SHEIN_INPUT_FILENAME
    python ebay_price_check.py "path/to/x.xlsx" # explicit file
```

to:

```
Usage:
    python ebay_price_check.py                  # picks SUBMITTED_DIR/SHEIN_OUTPUT_FILENAME (enriched)
    python ebay_price_check.py "path/to/enriched.xlsx"  # explicit enriched file
```

- [ ] **Step 4: Verify import + tests still pass**

```bash
python -c "import ebay_price_check; print('ok')"
python test_ebay_url_clean.py
python test_ebay_parse.py
python test_ebay_extract.py
python test_ebay_io.py
python test_ebay_shorten_query.py
```

Expected: `ok` and five `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add ebay_price_check.py
git commit -m "refactor(ebay-check): read/write the enriched file (not master); drop INPUT_FILENAME dep"
```

---

## Task 7: Update `setup_wizard.py` to require both filenames

**Files:**
- Modify: `setup_wizard.py`

- [ ] **Step 1: Rename input-filename field label and make it required**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\setup_wizard.py`. Find in `_step_paths()`:

```python
        rows = [
            ("输入表所在目录", "submitted_dir", True),
            ("输出根目录（留空 = 与输入目录相同；店铺名自动作为子文件夹）", "output_dir", True),
            ("指定输入文件名（可选；留空 = 处理目录下所有 .xlsx）", "input_filename", False),
        ]
```

Change to:

```python
        rows = [
            ("输入表所在目录", "submitted_dir", True),
            ("输出根目录（留空 = 与输入目录相同；店铺名自动作为子文件夹）", "output_dir", True),
            ("主表(输入)文件名（必填，例: 澳洲希音链接 (输入) - ZR.xlsx）", "input_filename", False),
            ("富表(输出)文件名（必填，例: 澳洲希音链接 (输出) - ZR.xlsx）", "output_filename", False),
        ]
```

- [ ] **Step 2: Add output_filename to the initial values dict**

Find the `self.values` dict initialization near the top of the class. Look for:

```python
            "input_filename": "",
```

Add right after it (still inside the dict):

```python
            "output_filename": "",
```

- [ ] **Step 3: Add validation for both fields in on_next()**

Find in `_step_paths()`:

```python
        def on_next():
            for k, var in self._entries.items():
                self.values[k] = var.get().strip()
            sd = Path(self.values["submitted_dir"])
            if not sd.is_dir():
                messagebox.showerror("路径错误", f"输入目录不存在:\n{sd}")
                return
            if not self.values["output_dir"]:
                # 默认：输出根 = 输入目录。跑起来后会在这下面自动建 <店铺名>/
                # 每个 xlsx 的每个 sheet 会成为一层子文件夹。
                self.values["output_dir"] = str(sd)
```

Insert validation for both filenames BEFORE the output_dir default assignment:

```python
        def on_next():
            for k, var in self._entries.items():
                self.values[k] = var.get().strip()
            sd = Path(self.values["submitted_dir"])
            if not sd.is_dir():
                messagebox.showerror("路径错误", f"输入目录不存在:\n{sd}")
                return
            if not self.values["input_filename"]:
                messagebox.showerror("配置缺失", "请填写'主表(输入)文件名'。")
                return
            if not self.values["output_filename"]:
                messagebox.showerror("配置缺失", "请填写'富表(输出)文件名'。")
                return
            if not self.values["output_dir"]:
                # 默认：输出根 = 输入目录。跑起来后会在这下面自动建 <店铺名>/
                # 每个 xlsx 的每个 sheet 会成为一层子文件夹。
                self.values["output_dir"] = str(sd)
```

- [ ] **Step 4: Wire output_filename to the config.env write**

Find the env-write dict in `_save_config()` or equivalent (search for `SHEIN_INPUT_FILENAME` in the file):

```python
        env_values = {
            "SHEIN_SUBMITTED_DIR": self.values["submitted_dir"],
            "SHEIN_OUTPUT_DIR": self.values["output_dir"],
        }
        if self.values["input_filename"]:
            env_values["SHEIN_INPUT_FILENAME"] = self.values["input_filename"]
```

Change the block to always write both (since both are now required):

```python
        env_values = {
            "SHEIN_SUBMITTED_DIR": self.values["submitted_dir"],
            "SHEIN_OUTPUT_DIR": self.values["output_dir"],
            "SHEIN_INPUT_FILENAME": self.values["input_filename"],
            "SHEIN_OUTPUT_FILENAME": self.values["output_filename"],
        }
```

(Remove the `if self.values["input_filename"]:` conditional entirely.)

Also, find the `_read_env_file` / config load logic and add `SHEIN_OUTPUT_FILENAME`. Look for:

```python
            "SHEIN_SUBMITTED_DIR": "submitted_dir",
            "SHEIN_OUTPUT_DIR": "output_dir",
            "SHEIN_INPUT_FILENAME": "input_filename",
```

Add:

```python
            "SHEIN_SUBMITTED_DIR": "submitted_dir",
            "SHEIN_OUTPUT_DIR": "output_dir",
            "SHEIN_INPUT_FILENAME": "input_filename",
            "SHEIN_OUTPUT_FILENAME": "output_filename",
```

- [ ] **Step 5: Verify existing setup wizard test still passes**

```bash
python test_setup_wizard.py
```

Expected: `ALL PASS` (test doesn't exercise the new field, but should not regress).

- [ ] **Step 6: Commit**

```bash
git add setup_wizard.py
git commit -m "feat(wizard): require both master (input) and enriched (output) filenames"
```

---

## Task 8: Version bump 0.2.1 → 0.3.0

**Files:**
- Modify: `version.py`
- Modify: `installer.iss`

- [ ] **Step 1: Bump version.py**

Change:
```python
VERSION = "0.2.1"
```
To:
```python
VERSION = "0.3.0"
```

- [ ] **Step 2: Bump installer.iss**

Change:
```
#define MyAppVersion "0.2.1"          ; Keep in sync with version.py
```
To:
```
#define MyAppVersion "0.3.0"          ; Keep in sync with version.py
```

- [ ] **Step 3: Commit**

```bash
git add version.py installer.iss
git commit -m "chore: bump 0.2.1 -> 0.3.0 for master/enriched split (breaking config)"
```

---

## Task 9: End-to-end manual smoke against real ZR files

**Files:** none new — verification against the operator's real data.

- [ ] **Step 1: Verify master file is set up correctly**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook
p = r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接 (输入) - ZR.xlsx'
wb = load_workbook(p, data_only=True)
for name in wb.sheetnames:
    ws = wb[name]
    print(f'=== Sheet: {name} ({ws.max_row}r x {ws.max_column}c) ===')
    print('Headers:', [ws.cell(1, c).value for c in range(1, ws.max_column+1)])
    y_rows = 0
    for r in range(2, ws.max_row+1):
        if str(ws.cell(r, 6).value or '').strip().upper() == 'Y':
            y_rows += 1
    print(f'F=Y rows: {y_rows}')
"
```

Expected: 6 headers matching MASTER_EXPECTED_HEADERS, `F=Y rows > 0`.

- [ ] **Step 2: Verify enriched file exists (or will be created)**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
p = Path(r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接 (输出) - ZR.xlsx')
print(f'exists? {p.exists()}')
if p.exists():
    from openpyxl import load_workbook
    wb = load_workbook(p)
    for name in wb.sheetnames:
        ws = wb[name]
        print(f'  {name}: {ws.max_row} rows (col A samples: {[ws.cell(r,1).value for r in range(1, min(ws.max_row+1, 5))]})')
"
```

If it doesn't exist yet, the first run will create it. If it exists, headers must match `EXPECTED_HEADERS`.

- [ ] **Step 3: Run the scrape**

```bash
python -X utf8 run_excel.py
```

Expect:
- Log lines showing master file opened, per-sheet count of F=Y rows
- Chrome opens, scrape runs (3-tab parallelism, ~10s/URL)
- Per-sheet: `Saved enriched progress to 澳洲希音链接 (输出) - ZR.xlsx`

- [ ] **Step 4: Verify enriched file was appended to (not master)**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
from openpyxl import load_workbook

master = Path(r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接 (输入) - ZR.xlsx')
enriched = Path(r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接 (输出) - ZR.xlsx')

for label, p in [('MASTER', master), ('ENRICHED', enriched)]:
    print(f'=== {label}: {p.name} ===')
    print(f'  mtime: {p.stat().st_mtime}')
    wb = load_workbook(p, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        print(f'  sheet {name!r}: {ws.max_row} rows, {ws.max_column} cols')
"
```

Expected:
- MASTER: mtime unchanged from before the run (script did not touch it).
- ENRICHED: mtime very recent; rows increased by the count of F=Y rows.

- [ ] **Step 5: Run ebay_price_check + verify**

```bash
python -X utf8 ebay_price_check.py
```

Then verify N-R columns filled on the newly-appended enriched rows:

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook
p = r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接 (输出) - ZR.xlsx'
wb = load_workbook(p, data_only=True)
ws = wb['ZR1']
print(f'seq  date        status    N搜索日期      O低价     Q高价')
for r in range(2, ws.max_row+1):
    seq = ws.cell(r, 1).value
    date = ws.cell(r, 6).value
    status = ws.cell(r, 7).value
    n = ws.cell(r, 14).value
    o = ws.cell(r, 15).value
    q = ws.cell(r, 17).value
    print(f'{seq!s:<4} {date!s:<11} {status!s:<9} {n!s:<12} {o!s:<9} {q!s}')
"
```

Expected: N/O/P/Q/R filled for rows where G=Done AND K non-empty.

- [ ] **Step 6: Verify master file mtime STILL unchanged**

Run Step 4's mtime check again. Master's mtime should still be from before Step 3. Enriched's mtime updated by BOTH steps 3 and 5.

- [ ] **Step 7: No commit for this task** — verification only. If issues found, escalate before proceeding to Task 10.

---

## Task 10: Rebuild installer + rollout

**Files:** none new — packaging step.

- [ ] **Step 1: Extract .build_key.txt from .env + generate key_store.py**

```powershell
Set-Location "C:\Users\ak\Desktop\Claude\shein-extract-au"
python -c "
with open('.env', encoding='utf-8') as f:
    for line in f:
        if line.startswith('ANTHROPIC_API_KEY='):
            with open('.build_key.txt', 'w', encoding='utf-8') as out:
                out.write(line.split('=', 1)[1].strip())
            break
"
python make_key_store.py
```

Expected: `wrote key_store.py (144 encoded chars)`.

- [ ] **Step 2: PyInstaller via Anaconda (bypass build.bat's cmd bug)**

```powershell
$env:PATH = "C:\Users\ak\Anaconda3;C:\Users\ak\Anaconda3\Scripts;C:\Users\ak\Anaconda3\Library\bin;" + $env:PATH
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist\SheinExtractAU.exe) { Remove-Item dist\SheinExtractAU.exe -Force }
python -m PyInstaller pyinstaller.spec --clean --noconfirm
```

Expected: `Build complete! The results are available in: ...\dist`.

- [ ] **Step 3: Inno Setup**

```powershell
$iscc = "C:\Program Files (x86)\Inno Setup 6\iscc.exe"
& $iscc installer.iss
```

Expected: `Successful compile ... Resulting Setup program filename is: ...\dist\SheinExtractAU-Setup-0.3.0.exe`.

- [ ] **Step 4: Cleanup + verify dist state**

```powershell
Remove-Item .build_key.txt, key_store.py -Force -ErrorAction SilentlyContinue
if (Test-Path "dist\SheinExtractAU-Setup-0.2.1.exe") { Remove-Item "dist\SheinExtractAU-Setup-0.2.1.exe" -Force }
Get-ChildItem dist | Select-Object Name, LastWriteTime | Format-Table -AutoSize
```

Expected: two files — `SheinExtractAU-Setup-0.3.0.exe` and `SheinExtractAU.exe`, both fresh.

- [ ] **Step 5: Push branch to origin (no merge — user decides when to merge to main)**

```bash
git push -u origin feat/master-enriched-split
```

- [ ] **Step 6: Communicate rollout note to operator**

Instructions to include when handing over:

```
装 0.3.0 后必须操作：
1. 卸载 0.2.1 (可选，0.3.0 会替换)
2. 双击 SheinExtractAU-Setup-0.3.0.exe 装
3. 首次跑会弹 setup_wizard —— 4 个字段全部填写:
   - 输入表所在目录
   - 输出根目录 (可留空)
   - 主表(输入)文件名: 你的主表 .xlsx 全名
   - 富表(输出)文件名: 你的富表 .xlsx 全名
4. wizard 结束后跑一遍 scrape 测试 —— 主表 mtime 应不变，富表被追加

BREAKING: 老单文件工作流不再支持。你的主表现在是"只读事实源"，脚本不动。
```

- [ ] **Step 7: This task ends with the installer artifact on the local machine + the branch pushed. No git commit needed — the installer isn't tracked.**

---

## Rollout

- Merge `feat/master-enriched-split` to `main` after Task 9 smoke passes.
- Distribute new installer to employees.
- Update MEMORY.md if the schema-split becomes a project-durable fact worth remembering across sessions.
