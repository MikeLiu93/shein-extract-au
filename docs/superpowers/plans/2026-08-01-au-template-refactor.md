# AU Template Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `shein-extract-au`'s dual-file input/output flow with a single Chinese-header template that acts as both input and primary output, port the AI-key bundling from the US project so employee installers stop falling back to rule-based titles, fix pacing to 3 s, and add 3-tab parallelism.

**Architecture:**
- `run_excel.py` reads/writes one Chinese-header template (12 cols A–L). Old English-header path is deleted.
- `shein_scraper.py` grows three pure helpers (`_filter_variants_by_declaration`, `_format_price_range`, `_format_stock_summary`), and its main loop is refactored into a per-URL worker driven by a `ThreadPoolExecutor(max_workers=3)`. Excel writes are guarded by a `threading.Lock`; the rate-limit counter is atomic.
- `make_key_store.py` + `build.bat` step ported verbatim from `shein-extract` (US) — same XOR/base64 scheme so `key_store.py` ships inside the compiled exe.
- Batch summary xlsx (`{store}-{minSeq}-{maxSeq}-{YYYYMMDD}.xlsx`) is retained untouched as the audit artifact.
- Reference spec: `docs/superpowers/specs/2026-08-01-au-template-refactor-design.md`.

**Tech Stack:** Python 3.10+, openpyxl, requests, websocket-client, PyInstaller, Inno Setup. Tests are plain-`assert` scripts (project convention — see `test_variant_merge.py`), runnable via `python test_<name>.py`.

**Working branch:** `feat/template-refactor` (already checked out at commit `b8bcf01`).

---

## File Structure

**New files:**
- `make_key_store.py` — build-time XOR/b64 encoder for the raw Anthropic API key.
- `key_store_template.py` — committed placeholder that documents what `key_store.py` (gitignored, generated) will look like.
- `docs/known-issues.md` — records the deferred item-2 variant-price problem.
- `test_template_io.py` — tests for the new template reader/writer functions (see Tasks 7–8).
- `test_variant_filter.py` — tests for the variant-filter helper (Task 4).
- `test_price_stock_format.py` — tests for the price-range and stock-summary helpers (Tasks 5–6).

**Modified files:**
- `shein_scraper.py`
  - Delete `INTER_URL_DELAY_MIN/MAX`, `LONG_PAUSE_EVERY/MIN/MAX`; rewrite `_inter_url_pause`.
  - Change `IMAGE_DOWNLOAD_WORKERS` from 8 to 4.
  - Add three new helpers listed above.
  - Refactor `scrape_shein()` main loop into a worker `_scrape_one_url(...)` + `ThreadPoolExecutor(max_workers=3)`; thread-safe rate-limit counter; add `variant_filter_list` param.
- `run_excel.py`
  - Delete old English-schema code paths.
  - Add `_read_pending_rows(ws)` and `_write_result_row(...)` for the Chinese schema.
  - Rewire `process_excel()` accordingly; pass `variant_filter_list` to `scrape_shein`.
- `build.bat`
  - Insert `python make_key_store.py` step before PyInstaller.
- `.gitignore`
  - Add `.build_key.txt` and `key_store.py`.
- `pyinstaller.spec`
  - Add `key_store` to hidden imports (safe no-op if the module isn't present in source mode).

---

## Task 1: Record deferred item 2 in `docs/known-issues.md`

**Files:**
- Create: `docs/known-issues.md`

- [ ] **Step 1: Create the file with the item-2 entry**

Write `C:\Users\ak\Desktop\Claude\shein-extract-au\docs\known-issues.md`:

```markdown
# Known Issues (deferred)

## Variants — per-variant price/stock may mis-associate

**Symptom:** in the batch xlsx or in `图片-<sku>/eBay上架描述.txt`, a variant
occasionally shows a wrong price, an empty stock, or a variant that doesn't
actually exist on the page. In the new template's `库存` column, the same
mismatch surfaces as a `-` where a number is expected.

**Suspected cause:** the mapping between `sku_prices` entries and their
attribute values inside `window.gbRawData.modules.saleAttr.multiLevelSaleAttribute.sku_list`
is fragile on some product layouts; the extraction JS in `shein_scraper.py::_JS`
sometimes hands the wrong attr row to the wrong SKU.

**Workaround (in place since 2026-08-01):** the new template's `E 变体` column
lets the operator pre-declare exactly which variants to list. When filled,
non-matching `sku_prices` entries are dropped before write, so mis-associations
outside the declared set do not appear.

**When to actually fix:** capture the failing URL + full `window.gbRawData`
JSON, then refine the `extractSkuPrices()` / attribute-mapping logic. Do not
attempt a speculative fix without a reproduction.
```

- [ ] **Step 2: Commit**

```powershell
git add docs/known-issues.md
git commit -m "docs: record item-2 variant mis-association as deferred known issue"
```

Expected: single-file commit, no other changes.

---

## Task 2: Port the AI-key packaging from `shein-extract` (item 1)

**Files:**
- Create: `make_key_store.py`
- Create: `key_store_template.py`
- Modify: `build.bat`
- Modify: `.gitignore`
- Modify: `pyinstaller.spec`

- [ ] **Step 1: Copy `make_key_store.py` verbatim from the US project**

Source: `C:\Users\ak\Desktop\Claude\shein-extract\make_key_store.py`.
Destination: `C:\Users\ak\Desktop\Claude\shein-extract-au\make_key_store.py`.

Copy the file byte-for-byte. Content (77 lines) — reproduced here so the plan is self-contained:

```python
"""
Build-time helper: read the raw Anthropic API key from .build_key.txt
(gitignored, owner-only) and write key_store.py (gitignored) with an
XOR + base64 obfuscated copy. PyInstaller then bundles key_store.py
into the .exe so employees never see the plaintext.

Usage:
    python make_key_store.py
    python make_key_store.py --key sk-ant-...    # override file lookup
"""

import argparse
import base64
import sys
from pathlib import Path

PASSPHRASE = "shein-extract-2026"
HERE = Path(__file__).resolve().parent
RAW_KEY_FILE = HERE / ".build_key.txt"
OUTPUT_FILE = HERE / "key_store.py"


def encode(raw_key: str) -> str:
    if not raw_key:
        return ""
    out = bytearray()
    for i, ch in enumerate(raw_key.encode("utf-8")):
        out.append(ch ^ ord(PASSPHRASE[i % len(PASSPHRASE)]))
    return base64.b64encode(bytes(out)).decode("ascii")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", help="Raw API key. If omitted, reads .build_key.txt")
    args = ap.parse_args()

    key = (args.key or "").strip()
    if not key:
        if not RAW_KEY_FILE.exists():
            print(f"[error] {RAW_KEY_FILE.name} not found.", file=sys.stderr)
            print("Create that file with the raw Anthropic API key on one line, "
                  "or pass --key sk-ant-...", file=sys.stderr)
            sys.exit(1)
        key = RAW_KEY_FILE.read_text(encoding="utf-8").strip()

    if not key.startswith("sk-ant-"):
        print(f"[error] Key doesn't look like an Anthropic key (must start "
              f"with 'sk-ant-'). Got: {key[:10]}...", file=sys.stderr)
        sys.exit(1)

    encoded = encode(key)

    OUTPUT_FILE.write_text(
        '"""Auto-generated by make_key_store.py — DO NOT EDIT, DO NOT COMMIT."""\n'
        f'ENCODED_KEY = "{encoded}"\n'
        f'PASSPHRASE = "{PASSPHRASE}"\n'
        '\n'
        'def get_obfuscated_key() -> str:\n'
        '    if not ENCODED_KEY:\n'
        '        return ""\n'
        '    import base64\n'
        '    try:\n'
        '        raw = base64.b64decode(ENCODED_KEY)\n'
        '        out = bytearray()\n'
        '        for i, b in enumerate(raw):\n'
        '            out.append(b ^ ord(PASSPHRASE[i % len(PASSPHRASE)]))\n'
        '        return out.decode("utf-8")\n'
        '    except Exception:\n'
        '        return ""\n',
        encoding="utf-8",
    )
    print(f"  wrote {OUTPUT_FILE.name} ({len(encoded)} encoded chars)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Create `key_store_template.py` (committed placeholder)**

Write `C:\Users\ak\Desktop\Claude\shein-extract-au\key_store_template.py`:

```python
"""
TEMPLATE — do not edit directly. The build script `make_key_store.py` reads
.build_key.txt (owner-only, gitignored) and writes the real
`key_store.py` (also gitignored) with the encoded key.

key_store.py is bundled into the .exe by PyInstaller so employees never see
the plaintext key.

Source-run developers can rely on the `.env` fallback in `shein_scraper.py`
instead — set ANTHROPIC_API_KEY there.
"""

ENCODED_KEY = ""
PASSPHRASE = "shein-extract-2026"


# Will be replaced by make_key_store.py with the actual encoded value.
def get_obfuscated_key() -> str:
    return ""
```

- [ ] **Step 3: Add `.build_key.txt` and `key_store.py` to `.gitignore`**

Append to the end of `C:\Users\ak\Desktop\Claude\shein-extract-au\.gitignore`:

```
# AI key bundling (never commit)
.build_key.txt
key_store.py
```

- [ ] **Step 4: Insert key-generation step in `build.bat`**

Modify `C:\Users\ak\Desktop\Claude\shein-extract-au\build.bat`. Find the block that starts with:

```
echo.
echo ============================================================
echo  Step 1/2: PyInstaller - build SheinExtractAU.exe
```

Replace with:

```
echo.
echo ============================================================
echo  Step 1/3: Generate key_store.py from .build_key.txt
echo ============================================================
if not exist .build_key.txt (
    echo [ERROR] .build_key.txt not found. Create it with the raw API key.
    exit /b 1
)
python make_key_store.py
if errorlevel 1 (
    echo [ERROR] make_key_store.py failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 2/3: PyInstaller - build SheinExtractAU.exe
echo ============================================================
```

Also update the block header near the bottom from `Step 2/2: Inno Setup` to `Step 3/3: Inno Setup`.

- [ ] **Step 5: Add `key_store` to `pyinstaller.spec` hidden imports**

Modify `C:\Users\ak\Desktop\Claude\shein-extract-au\pyinstaller.spec`. Find the `hiddenimports=[` list. After the `'auth',` line, add:

```python
        'auth',
        'version',
        # Optional — only present after make_key_store.py ran
        'key_store',
```

(Do not remove `'version',` — insert the new line after it. If `key_store.py` isn't present when PyInstaller runs, this hidden import is a warning, not an error; it becomes real once `make_key_store.py` runs first.)

- [ ] **Step 6: Sanity-check the encoder works locally**

Run:

```powershell
"sk-ant-fake-key-for-encoder-check" | Out-File -Encoding utf8 .build_key.txt
python make_key_store.py
```

Expected output line: `wrote key_store.py (<N> encoded chars)`.

Then round-trip:

```powershell
python -c "from key_store import get_obfuscated_key; print(get_obfuscated_key())"
```

Expected: `sk-ant-fake-key-for-encoder-check`.

Then delete both files (they must not enter the commit):

```powershell
Remove-Item .build_key.txt, key_store.py -Force
```

- [ ] **Step 7: Commit**

```powershell
git add make_key_store.py key_store_template.py .gitignore build.bat pyinstaller.spec
git commit -m "build: port AI-key bundling from shein-extract (fixes silent AI-title fallback)"
```

Expected: 5 files changed, no `key_store.py` or `.build_key.txt` in the diff.

---

## Task 3: Fixed 3-second pacing + smaller image pool (items 3 & 4 prep)

**Files:**
- Modify: `shein_scraper.py:78-108` (constants + `_inter_url_pause`)
- Modify: `shein_scraper.py:93` (`IMAGE_DOWNLOAD_WORKERS`)

- [ ] **Step 1: Replace pacing constants with a single fixed delay**

In `C:\Users\ak\Desktop\Claude\shein-extract-au\shein_scraper.py`, delete lines defining `INTER_URL_DELAY_MIN`, `INTER_URL_DELAY_MAX`, `LONG_PAUSE_EVERY`, `LONG_PAUSE_MIN`, `LONG_PAUSE_MAX` (currently around lines 75–81). Add one new constant in their place:

```python
# Fixed pause between URL batches. Under parallelism this is the gap between
# batches of MAX_WORKERS, not between individual URLs. 3s was chosen after
# discussion — enough for Chrome tab cleanup and image processing without
# feeling artificial.
INTER_URL_DELAY_SEC    = 3
```

- [ ] **Step 2: Rewrite `_inter_url_pause`**

Replace the current body of `_inter_url_pause` with:

```python
def _inter_url_pause(i: int, total: int) -> None:
    """Sleep between URL batches. Skip after last item."""
    if i >= total:
        return
    print(f"  [节奏] 间隔 {INTER_URL_DELAY_SEC}s")
    time.sleep(INTER_URL_DELAY_SEC)
```

**Do NOT remove `import random`** — the OOPS backoff block (around line 2604) still uses `random.uniform(OOPS_RETRY_BACKOFF_MIN, OOPS_RETRY_BACKOFF_MAX)`.

- [ ] **Step 3: Reduce `IMAGE_DOWNLOAD_WORKERS` from 8 to 4**

In `shein_scraper.py`, change:

```python
IMAGE_DOWNLOAD_WORKERS = 8        # 并行下载线程数
```

to:

```python
IMAGE_DOWNLOAD_WORKERS = 4        # 并行下载线程数（3并发×4=12 sockets 总量可控）
```

- [ ] **Step 4: Quick smoke — module still imports**

Run:

```powershell
python -c "import shein_scraper; print(shein_scraper.INTER_URL_DELAY_SEC, shein_scraper.IMAGE_DOWNLOAD_WORKERS)"
```

Expected output: `3 4`.

- [ ] **Step 5: Commit**

```powershell
git add shein_scraper.py
git commit -m "perf(scraper): fixed 3s inter-URL pause; image pool 8->4 for parallelism"
```

---

## Task 4: Variant-filter helper — TDD

**Files:**
- Create: `test_variant_filter.py`
- Modify: `shein_scraper.py` (append helper near other pure functions, e.g. after `_split_variations_for_excel`)

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_variant_filter.py`:

```python
"""Tests for _filter_variants_by_declaration — the E-column filter that
restricts sku_prices + variations to only the user's declared variants."""
from shein_scraper import _filter_variants_by_declaration


def _make_sku(sku_code, color, size, price=9.99, stock=10):
    return {
        "sku_code": sku_code,
        "attrs": {"Color": color, "Size": size},
        "sale_price": price,
        "retail_price": price,
        "stock": stock,
    }


SAMPLE_SKUS = [
    _make_sku("A1", "Black", "M", 9.99, 12),
    _make_sku("A2", "Black", "L", 9.99, 0),
    _make_sku("A3", "Red",   "M", 10.99, 3),
    _make_sku("A4", "Red",   "L", 10.99, 20),
    _make_sku("A5", "Blue",  "M", 9.99, 5),
]
SAMPLE_VARS = {"Color": ["Black", "Red", "Blue"], "Size": ["M", "L"]}


def test_empty_declaration_returns_everything():
    kept_skus, kept_vars, unknown = _filter_variants_by_declaration(
        "", SAMPLE_SKUS, SAMPLE_VARS
    )
    assert kept_skus == SAMPLE_SKUS
    assert kept_vars == SAMPLE_VARS
    assert unknown == []


def test_two_groups_cartesian():
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Red / M, L", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A2", "A3", "A4"], codes
    assert vars_ == {"Color": ["Black", "Red"], "Size": ["M", "L"]}
    assert unknown == []


def test_single_group_flat_allowlist():
    # No "/" → allow-list matched against any attribute value.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Red", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A2", "A3", "A4"], codes
    assert vars_["Color"] == ["Black", "Red"]
    assert unknown == []


def test_case_and_whitespace_insensitive():
    kept, vars_, unknown = _filter_variants_by_declaration(
        "  black ,RED /  m ", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A3"], codes
    assert unknown == []


def test_unknown_value_but_partial_match():
    # "Green" doesn't exist; "Black" does. Keep Black rows, report Green.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Green / M", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1"], codes
    assert "Green" in unknown, unknown


def test_zero_match_returns_empty_and_flags():
    # Nothing matches — caller (run_excel) will surface as Failed.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Purple / XXL", SAMPLE_SKUS, SAMPLE_VARS
    )
    assert kept == []
    assert vars_ == {}
    assert set(unknown) >= {"Purple", "XXL"}


if __name__ == "__main__":
    test_empty_declaration_returns_everything()
    test_two_groups_cartesian()
    test_single_group_flat_allowlist()
    test_case_and_whitespace_insensitive()
    test_unknown_value_but_partial_match()
    test_zero_match_returns_empty_and_flags()
    print("ALL PASS")
```

- [ ] **Step 2: Run tests — expect ImportError**

```powershell
python test_variant_filter.py
```

Expected: `ImportError: cannot import name '_filter_variants_by_declaration' from 'shein_scraper'`.

- [ ] **Step 3: Implement the helper in `shein_scraper.py`**

Append after `_split_variations_for_excel` (around line 1422). Add:

```python
def _filter_variants_by_declaration(
    declaration: str,
    sku_prices: list,
    variations: dict,
) -> tuple[list, dict, list]:
    """Restrict sku_prices + variations to variants declared by the user.

    Grammar of `declaration`:
      - Empty / whitespace → no filter; return inputs unchanged.
      - No "/"             → flat allow-list; a SKU is kept if any of its
                             attribute values matches (case/whitespace-insensitive).
      - Contains "/"       → groups separated by "/", values within each group
                             comma-separated. A SKU is kept only if EVERY
                             non-empty group has at least one value equal to
                             one of that SKU's attribute values.

    Returns (kept_sku_prices, filtered_variations_dict, unknown_values_list).
    `unknown_values_list` contains declared values that never matched any
    scraped variant — caller logs a warning. When declaration was given
    but nothing matched, kept is [] and filtered_variations is {}.
    """
    decl = (declaration or "").strip()
    if not decl:
        return sku_prices, variations, []

    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "")).strip().lower()

    groups = [
        [_norm(v) for v in grp.split(",") if _norm(v)]
        for grp in decl.split("/")
    ]
    groups = [g for g in groups if g]  # drop empty groups
    if not groups:
        return sku_prices, variations, []

    all_declared = {v for g in groups for v in g}
    seen_values: set[str] = set()

    def _sku_matches(sku: dict) -> bool:
        attr_values = {_norm(v) for v in (sku.get("attrs") or {}).values() if v}
        if len(groups) == 1:
            # Flat allow-list: any attribute value overlaps the group.
            hits = attr_values & set(groups[0])
            seen_values.update(hits)
            return bool(hits)
        # Multi-group: every group must overlap this SKU's attribute values.
        ok = True
        for g in groups:
            hits = attr_values & set(g)
            if not hits:
                ok = False
            else:
                seen_values.update(hits)
        return ok

    kept = [s for s in sku_prices if _sku_matches(s)]
    unknown = sorted(all_declared - seen_values)

    if not kept:
        return [], {}, sorted(all_declared)

    # Rebuild variations dict from kept SKUs so downstream (title, txt, L col)
    # sees only declared values in original scraper casing.
    filtered_vars: dict[str, list] = {}
    for sp in kept:
        for k, v in (sp.get("attrs") or {}).items():
            if not v:
                continue
            filtered_vars.setdefault(k, [])
            if v not in filtered_vars[k]:
                filtered_vars[k].append(v)
    # Preserve top-level variations keys that were flat (no per-SKU attrs) —
    # rare but safe: keep any original key not covered by kept-SKU attrs, if
    # its values overlap the declaration.
    for k, vals in (variations or {}).items():
        if k in filtered_vars:
            continue
        if not isinstance(vals, list):
            continue
        keep = [v for v in vals if _norm(v) in all_declared]
        if keep:
            filtered_vars[k] = keep

    return kept, filtered_vars, unknown
```

Also confirm `re` and `list`/`dict`/`tuple`/`set` are already imported at the top of the file (they are — `re` is imported on line ~42; the typing generics are Python 3.9+ builtin).

- [ ] **Step 4: Run tests — expect PASS**

```powershell
python test_variant_filter.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_variant_filter.py shein_scraper.py
git commit -m "feat(scraper): add _filter_variants_by_declaration for E-column filter"
```

---

## Task 5: Price-range formatter — TDD

**Files:**
- Create (or extend): `test_price_stock_format.py`
- Modify: `shein_scraper.py` (append near `_split_variations_for_excel`)

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_price_stock_format.py`:

```python
"""Tests for _format_price_range (H column '希音价格') and _format_stock_summary
(L column '库存') — Chinese-header template output columns."""
from shein_scraper import _format_price_range, _format_stock_summary


# ── _format_price_range ─────────────────────────────────────────────────────

def test_empty_returns_empty_string():
    assert _format_price_range([]) == ""


def test_single_price_no_range():
    skus = [{"sale_price": 9.99}]
    assert _format_price_range(skus) == "$9.99"


def test_all_same_price_no_range():
    skus = [{"sale_price": 9.99}, {"sale_price": 9.99}, {"sale_price": 9.99}]
    assert _format_price_range(skus) == "$9.99"


def test_differing_prices_range():
    skus = [{"sale_price": 9.99}, {"sale_price": 14.99}, {"sale_price": 12.50}]
    assert _format_price_range(skus) == "$9.99–$14.99"


def test_ignores_none_prices():
    skus = [{"sale_price": None}, {"sale_price": 9.99}]
    assert _format_price_range(skus) == "$9.99"


def test_only_nones_returns_empty():
    skus = [{"sale_price": None}, {"sale_price": None}]
    assert _format_price_range(skus) == ""


# ── _format_stock_summary ───────────────────────────────────────────────────

def test_stock_empty_returns_empty():
    assert _format_stock_summary([]) == ""


def test_stock_single_variant_no_prefix():
    skus = [{"attrs": {"Color": "Black"}, "stock": 12}]
    assert _format_stock_summary(skus) == "Black: 12"


def test_stock_multi_variant_slash_separated():
    skus = [
        {"attrs": {"Color": "Black", "Size": "M"}, "stock": 12},
        {"attrs": {"Color": "Black", "Size": "L"}, "stock": 0},
        {"attrs": {"Color": "Red",   "Size": "M"}, "stock": 3},
    ]
    out = _format_stock_summary(skus)
    assert out == "Black-M: 12 / Black-L: 缺货 / Red-M: 少货 3", out


def test_stock_low_stock_threshold_15():
    # LOW_STOCK_THRESHOLD is 15 in shein_scraper.
    skus = [
        {"attrs": {"Size": "M"}, "stock": 15},   # exactly at threshold → 少货
        {"attrs": {"Size": "L"}, "stock": 16},   # above → plain number
    ]
    out = _format_stock_summary(skus)
    assert out == "M: 少货 15 / L: 16", out


def test_stock_missing_attrs_uses_sku_code():
    skus = [{"sku_code": "ABC123", "attrs": {}, "stock": 5}]
    out = _format_stock_summary(skus)
    assert out == "ABC123: 少货 5", out


if __name__ == "__main__":
    test_empty_returns_empty_string()
    test_single_price_no_range()
    test_all_same_price_no_range()
    test_differing_prices_range()
    test_ignores_none_prices()
    test_only_nones_returns_empty()
    test_stock_empty_returns_empty()
    test_stock_single_variant_no_prefix()
    test_stock_multi_variant_slash_separated()
    test_stock_low_stock_threshold_15()
    test_stock_missing_attrs_uses_sku_code()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ImportError**

```powershell
python test_price_stock_format.py
```

Expected: `ImportError: cannot import name '_format_price_range' from 'shein_scraper'`.

- [ ] **Step 3: Implement `_format_price_range`**

Append to `shein_scraper.py` near the other formatters (after `_filter_variants_by_declaration`):

```python
def _format_price_range(sku_prices: list) -> str:
    """Return '$X.XX' if all SKU sale_prices are equal, else '$LOW–$HIGH'.
    Uses an en-dash (U+2013). Ignores None prices. Returns '' if none valid."""
    prices = []
    for sp in sku_prices or []:
        p = sp.get("sale_price")
        if p is None:
            continue
        try:
            prices.append(float(p))
        except (TypeError, ValueError):
            continue
    if not prices:
        return ""
    lo, hi = min(prices), max(prices)
    if lo == hi:
        return f"${lo:.2f}"
    return f"${lo:.2f}–${hi:.2f}"
```

- [ ] **Step 4: Run price tests — 6 pass, stock tests still fail**

```powershell
python test_price_stock_format.py
```

Expected: fails on the first stock test with `ImportError` (or `AttributeError` if the import block imports both — the file imports both at top, so it'll still be `ImportError`). This is a single test file; either implement stock in the next step and re-run, or split. Continue to Task 6 immediately.

---

## Task 6: Stock-summary formatter — TDD (finishes Task 5's file)

**Files:**
- Modify: `shein_scraper.py` (append `_format_stock_summary`)

- [ ] **Step 1: Implement `_format_stock_summary`**

Append to `shein_scraper.py` immediately after `_format_price_range`:

```python
def _format_stock_summary(sku_prices: list) -> str:
    """One-line stock summary for the L column ('库存').

    Format: '<label>: <count>' entries joined by ' / '. Label is a hyphenated
    join of attribute values in insertion order (e.g. 'Black-M'); when the
    SKU has no attributes, sku_code is used. Count is 缺货 for 0, '少货 N'
    for 1..LOW_STOCK_THRESHOLD, and the plain number above that. When there
    is exactly one SKU with attribute values, the label is still emitted so
    the operator can see which variant the stock refers to.
    """
    def _label(sp: dict) -> str:
        vals = [str(v).strip() for v in (sp.get("attrs") or {}).values() if v]
        if vals:
            return "-".join(vals)
        return str(sp.get("sku_code") or "").strip() or "?"

    def _count(sp: dict) -> str:
        try:
            stk = int(sp.get("stock") or 0)
        except (TypeError, ValueError):
            stk = 0
        if stk == 0:
            return "缺货"
        if stk <= LOW_STOCK_THRESHOLD:
            return f"少货 {stk}"
        return str(stk)

    parts = [f"{_label(sp)}: {_count(sp)}" for sp in (sku_prices or [])]
    return " / ".join(parts)
```

- [ ] **Step 2: Run all format tests — expect PASS**

```powershell
python test_price_stock_format.py
```

Expected: `ALL PASS`.

- [ ] **Step 3: Commit**

```powershell
git add test_price_stock_format.py shein_scraper.py
git commit -m "feat(scraper): add _format_price_range + _format_stock_summary for template output"
```

---

## Task 7: Template reader (Chinese schema) — TDD

**Files:**
- Create: `test_template_io.py`
- Modify: `run_excel.py` (add `_read_pending_rows` function)

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_template_io.py`. This uses an in-memory Workbook — no fixture files:

```python
"""Tests for the Chinese-header template reader/writer in run_excel.py."""
from openpyxl import Workbook

from run_excel import _read_pending_rows, _write_result_row


CN_HEADERS = ["编号", "链接", "原价", "运费", "变体",
              "日期", "状态", "希音价格", "希音标题",
              "eBay标题", "eBay价格", "库存"]


def _blank_ws(rows: list[list]):
    wb = Workbook()
    ws = wb.active
    ws.title = "ZR1"
    for ci, h in enumerate(CN_HEADERS, 1):
        ws.cell(1, ci).value = h
    for ri, r in enumerate(rows, 2):
        for ci, v in enumerate(r, 1):
            ws.cell(ri, ci).value = v
    return ws


# ── _read_pending_rows ──────────────────────────────────────────────────────

def test_reads_only_rows_ready_and_not_done():
    # Row 2: ready. Row 3: has date → skip. Row 4: no URL → skip.
    ws = _blank_ws([
        [1, "http://u1", 10.0, 7.95, "",  None,        None, None, None, None, None, None],
        [2, "http://u2", 12.0, 7.95, "",  "2026-07-30", "Done", None, None, None, None, None],
        [3, None,        15.0, 7.95, "",  None,        None, None, None, None, None, None],
    ])
    pending = _read_pending_rows(ws)
    assert len(pending) == 1
    p = pending[0]
    assert p["row"] == 2
    assert p["seq"] == 1
    assert p["url"] == "http://u1"
    assert p["price"] == 10.0
    assert p["shipping"] == 7.95
    assert p["variant_filter"] == ""


def test_blank_price_and_shipping_kept_as_none():
    ws = _blank_ws([
        [1, "http://u1", None, None, "", None, None, None, None, None, None, None],
    ])
    pending = _read_pending_rows(ws)
    assert len(pending) == 1
    p = pending[0]
    assert p["price"] is None
    assert p["shipping"] is None


def test_variant_filter_captured():
    ws = _blank_ws([
        [1, "http://u1", 10.0, 7.95, "Black, Red / M, L", None, None, None, None, None, None, None],
    ])
    pending = _read_pending_rows(ws)
    assert pending[0]["variant_filter"] == "Black, Red / M, L"


def test_rejects_non_template_sheet():
    # A sheet missing the Chinese '链接' header in col B is not our schema.
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "Seq"
    ws.cell(1, 2).value = "Website"  # English → reject
    ws.cell(2, 1).value = 1
    ws.cell(2, 2).value = "http://u"
    pending = _read_pending_rows(ws)
    assert pending == []  # empty list = "not our schema, skip"


# ── _write_result_row ───────────────────────────────────────────────────────

def test_write_populates_result_columns():
    ws = _blank_ws([
        [1, "http://u1", 10.0, 7.95, "", None, None, None, None, None, None, None],
    ])
    _write_result_row(
        ws, row=2,
        date="2026-08-01",
        status="Done",
        web_price="$9.99–$14.99",
        shein_title="Foo Product",
        ebay_title="NEW Foo Product Adjustable Waterproof",
        ebay_price=27.95,
        stock="Black-M: 12 / Black-L: 缺货",
    )
    assert ws.cell(2, 6).value == "2026-08-01"
    assert ws.cell(2, 7).value == "Done"
    assert ws.cell(2, 8).value == "$9.99–$14.99"
    assert ws.cell(2, 9).value == "Foo Product"
    assert ws.cell(2, 10).value == "NEW Foo Product Adjustable Waterproof"
    assert ws.cell(2, 11).value == 27.95
    assert ws.cell(2, 12).value == "Black-M: 12 / Black-L: 缺货"


def test_write_only_date_status_on_failure():
    ws = _blank_ws([
        [1, "http://u1", 10.0, 7.95, "", None, None, None, None, None, None, None],
    ])
    _write_result_row(ws, row=2, date="2026-08-01", status="Failed")
    assert ws.cell(2, 6).value == "2026-08-01"
    assert ws.cell(2, 7).value == "Failed"
    # Result cols untouched
    assert ws.cell(2, 8).value is None
    assert ws.cell(2, 9).value is None
    assert ws.cell(2, 10).value is None
    assert ws.cell(2, 11).value is None
    assert ws.cell(2, 12).value is None


if __name__ == "__main__":
    test_reads_only_rows_ready_and_not_done()
    test_blank_price_and_shipping_kept_as_none()
    test_variant_filter_captured()
    test_rejects_non_template_sheet()
    test_write_populates_result_columns()
    test_write_only_date_status_on_failure()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ImportError**

```powershell
python test_template_io.py
```

Expected: `ImportError: cannot import name '_read_pending_rows' from 'run_excel'`.

- [ ] **Step 3: Add `_read_pending_rows` to `run_excel.py`**

Insert near the top of `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`, right after the imports (line ~35):

```python
# ── New Chinese-header template schema ────────────────────────────────────────
# Cols A–L. See docs/superpowers/specs/2026-08-01-au-template-refactor-design.md §1.
COL_SEQ, COL_URL, COL_PRICE, COL_SHIPPING, COL_VARIANT_FILTER = 1, 2, 3, 4, 5
COL_DATE, COL_STATUS, COL_WEB_PRICE, COL_SHEIN_TITLE = 6, 7, 8, 9
COL_EBAY_TITLE, COL_EBAY_PRICE, COL_STOCK = 10, 11, 12

EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体",
                    "日期", "状态", "希音价格", "希音标题",
                    "eBay标题", "eBay价格", "库存"]


def _sheet_matches_template(ws) -> bool:
    """The sheet must have '链接' in col B row 1 to be treated as the new schema."""
    return str(ws.cell(1, COL_URL).value or "").strip() == "链接"


def _read_pending_rows(ws) -> list[dict]:
    """Return [{row, seq, url, price, shipping, variant_filter}, ...] for
    rows that have 链接 filled and 日期/状态 both empty. Non-template sheets
    return []. Empty 原价/运费 come through as None (scraper fallback)."""
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        url = ws.cell(r, COL_URL).value
        date_v = ws.cell(r, COL_DATE).value
        status_v = ws.cell(r, COL_STATUS).value
        if not url or date_v or status_v:
            continue
        if seq is None:
            logger.info("  row %d: skip (no 编号)", r)
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
            "seq": int(seq),
            "url": str(url).strip(),
            "price": price,
            "shipping": shipping,
            "variant_filter": variant_filter,
        })
    return pending


def _write_result_row(
    ws, row: int, date: str, status: str,
    web_price: "str | None" = None,
    shein_title: "str | None" = None,
    ebay_title: "str | None" = None,
    ebay_price: "float | None" = None,
    stock: "str | None" = None,
) -> None:
    """Write result columns F–L. Optional cols default to None (skip write).
    Always writes 日期 (F) and 状态 (G)."""
    ws.cell(row, COL_DATE).value = date
    ws.cell(row, COL_STATUS).value = status
    if web_price is not None:
        ws.cell(row, COL_WEB_PRICE).value = web_price
    if shein_title is not None:
        ws.cell(row, COL_SHEIN_TITLE).value = shein_title
    if ebay_title is not None:
        ws.cell(row, COL_EBAY_TITLE).value = ebay_title
    if ebay_price is not None:
        ws.cell(row, COL_EBAY_PRICE).value = ebay_price
    if stock is not None:
        ws.cell(row, COL_STOCK).value = stock
```

- [ ] **Step 4: Run — expect PASS**

```powershell
python test_template_io.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_template_io.py run_excel.py
git commit -m "feat(run_excel): add Chinese-schema template reader/writer helpers"
```

---

## Task 8: Wire the new schema into `process_excel()` — remove old English path

**Files:**
- Modify: `run_excel.py` (rewrite `process_excel`, remove header-normalization block, delete old pending loop)

**Note:** This is where the old English format is deleted. There is no compatibility branch. Rows on old-format sheets are silently skipped (they return `[]` from `_read_pending_rows`).

- [ ] **Step 1: Rewrite `process_excel()`**

Open `C:\Users\ak\Desktop\Claude\shein-extract-au\run_excel.py`. Delete the body of `process_excel` (from `logger.info("Opening: %s"...` down to the final `logger.info("Done: %s", xlsx_path.name)`) and replace with:

```python
def process_excel(xlsx_path: Path) -> None:
    """Process all worksheets in `xlsx_path` using the Chinese-header template
    schema. Non-template sheets are skipped. See design spec §1 for columns."""
    logger.info("Opening: %s", xlsx_path.name)
    wb = load_workbook(xlsx_path)

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        store = ws_name.strip()
        pending = _read_pending_rows(ws)
        if not pending:
            logger.info("  Sheet '%s': no pending rows (or not template schema)", store)
            continue

        logger.info("  Sheet '%s': %d pending row(s): seq %s",
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

        # Fan results back into the template rows
        for p in pending:
            row = p["row"]
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
                _write_result_row(
                    ws, row=row,
                    date=today, status="Done",
                    web_price=rec.get("web_price_display"),
                    shein_title=rec.get("original_title") or rec.get("title"),
                    ebay_title=rec.get("ebay_title"),
                    ebay_price=rec.get("ebay_price"),
                    stock=rec.get("stock_summary"),
                )
                logger.info("    seq %d → Done", seq)
            elif rec and rec.get("status") == "DELISTED":
                _write_result_row(ws, row=row, date=today, status="Delisted")
                logger.info("    seq %d → Delisted", seq)
            else:
                detail = rec.get("status", "") if rec else ""
                if is_bad_data:
                    detail = "no data loaded"
                _write_result_row(ws, row=row, date=today, status="Failed")
                logger.info("    seq %d → Failed %s", seq,
                            f"({detail})" if detail else "")

        safe_save(wb, xlsx_path)
        logger.info("  Saved progress to %s", xlsx_path.name)

    wb.close()
    logger.info("Done: %s", xlsx_path.name)
```

- [ ] **Step 2: Remove old-schema code and `--no-price-gate` flag**

In the same file, update `main()`. Find:

```python
parser.add_argument("--no-price-gate", action="store_true",
                    help="处理 C 列 Price 为空的行(用网页价,不做覆盖);"
                         "默认要求 Price 填写才跑")
```

Delete that entire `add_argument` block. Then update the call:

```python
for f in files:
    logger.info("=" * 60)
    try:
        process_excel(f, gate_price=not args.no_price_gate)
    except Exception as e:
```

Change `process_excel(f, gate_price=not args.no_price_gate)` to `process_excel(f)`.

- [ ] **Step 3: Quick import check**

```powershell
python -c "import run_excel; print('ok')"
```

Expected: `ok`. If AttributeError on `_read_pending_rows` (order-of-definition), move the helper block above `process_excel`.

- [ ] **Step 4: Re-run the template IO tests — still PASS**

```powershell
python test_template_io.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add run_excel.py
git commit -m "refactor(run_excel): rewrite process_excel for Chinese-schema template; drop old English input path"
```

---

## Task 9: Extend `scrape_shein()` signature — accept `shipping_list` + `variant_filter_list` + emit `web_price_display` / `stock_summary` (single-threaded first)

**Files:**
- Modify: `shein_scraper.py` (extend `scrape_shein` signature; use inputs; still single-threaded — parallelism lands in Task 10)

This intermediate step delivers the record-shape changes so the run_excel wiring from Task 8 works end-to-end **before** we introduce concurrency. That keeps the concurrency change isolated.

- [ ] **Step 1: Extend signature and per-URL body**

In `shein_scraper.py`, change the `scrape_shein` signature:

```python
def scrape_shein(urls, output="shein_products.xlsx", start_seq=1, seq_list=None,
                 price_list=None, shipping_list=None, variant_filter_list=None):
```

Add to the docstring after `price_list`:

```
    shipping_list        : list[float|None] | None
                           每个 URL 对应的手动运费。非 None 时覆盖扫描运费。
    variant_filter_list  : list[str] | None
                           每个 URL 对应的变体过滤声明（模板 E 列）。
                           空串或 None 表示不过滤。
```

- [ ] **Step 2: Use the new params inside the main loop**

In the per-URL block of the loop (around line 2570+), after the existing `override_price` handling (line ~2708), add analogous shipping override + variant filter + summary generation. Find:

```python
shipping = _calc_shipping(data)
price    = data.get("price") or 0.0
ebay     = _ebay_listing_price(price, shipping)
```

Replace with:

```python
# Shipping override: template D column wins over scraped shipping.
override_shipping = (shipping_list[i - 1] if shipping_list else None)
if override_shipping is not None:
    shipping = float(override_shipping)
    ship_note = f"template D=${shipping:.2f} (override)"
else:
    shipping = _calc_shipping(data)
    ship_note = None  # will be overwritten below if not set here
price = data.get("price") or 0.0
ebay  = _ebay_listing_price(price, shipping)
```

Then find the block that computes the original `ship_note` (the `if data.get("unconditional_free"): ship_note = "unconditional FREE shipping"` branches). Wrap that entire block in `if ship_note is None:` so the override note isn't clobbered. Concretely, change:

```python
thresh   = data.get("free_threshold")
if data.get("unconditional_free"):
    ship_note = "unconditional FREE shipping"
elif thresh is not None:
    ...
```

to:

```python
thresh = data.get("free_threshold")
if ship_note is None:
    if data.get("unconditional_free"):
        ship_note = "unconditional FREE shipping"
    elif thresh is not None:
        ship_note = (
            f"threshold AU${thresh:.2f} — price AU${price:.2f} "
            + ("≥ threshold → FREE" if price >= thresh
               else f"< threshold → AU${DEFAULT_SHIPPING_FEE}")
        )
    elif not (data.get("shipping_raw") or "").strip():
        ship_note = "no shipping info on page → assumed FREE"
    else:
        ship_note = f"shipping text present but no threshold → AU${DEFAULT_SHIPPING_FEE}"
```

- [ ] **Step 3: Apply variant filter after `_merge_main_sale_attr_colors`**

Immediately after this existing line:

```python
data["variations"] = _merge_main_sale_attr_colors(
    data.get("variations") or {}, data.get("main_sale_attrs") or [])
```

Insert:

```python
# Variant filter: E-column declaration restricts sku_prices + variations.
vf_decl = (variant_filter_list[i - 1] if variant_filter_list else "")
if vf_decl:
    _kept, _filt_vars, _unknown = _filter_variants_by_declaration(
        vf_decl, data.get("sku_prices") or [], data.get("variations") or {}
    )
    if _unknown:
        print(f"  [变体过滤] 声明中未匹配到: {_unknown}")
    if not _kept and (data.get("sku_prices") or []):
        # Zero matches on a product that DID have variants — mark failed.
        rec["status"] = f"ERROR: variant_filter_no_match ({vf_decl!r})"
        records.append(rec)
        if tab_id is not None:
            _close_tab(CDP_PORT, tab_id)
        _inter_url_pause(i, len(urls))
        continue
    data["sku_prices"] = _kept
    if _filt_vars:
        data["variations"] = _filt_vars
```

- [ ] **Step 4: Emit `web_price_display` and `stock_summary` on the record**

Find the `rec.update({...})` block right after the shipping calc. Add three new keys to it (right after `"web_price": web_price,`):

```python
    "web_price":         web_price,
    "web_price_display": _format_price_range(data.get("sku_prices") or []) or (
                          f"${float(web_price):.2f}" if web_price is not None else ""),
    "stock_summary":     _format_stock_summary(data.get("sku_prices") or []),
    "shipping":          shipping,
```

- [ ] **Step 5: Quick smoke — module still imports and existing tests pass**

```powershell
python -c "import shein_scraper; import inspect; print(inspect.signature(shein_scraper.scrape_shein))"
```

Expected output includes `shipping_list=None, variant_filter_list=None`.

```powershell
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
```

Expected: each prints `ALL PASS`.

- [ ] **Step 6: Commit**

```powershell
git add shein_scraper.py
git commit -m "feat(scraper): scrape_shein accepts shipping_list + variant_filter_list; emits web_price_display / stock_summary"
```

---

## Task 10: 3-tab parallelism — refactor `scrape_shein` into ThreadPoolExecutor

**Files:**
- Modify: `shein_scraper.py` (extract per-URL body into `_scrape_one_url`; wrap dispatch in ThreadPoolExecutor; thread-safe rate-limit counter)

**Note:** This is the largest task. Extract the entire per-URL body from `scrape_shein`'s `for i, url in enumerate(urls, 1):` loop into a standalone function. The refactor is mechanical but touches ~200 lines. Do it in one go; the diff is verified by running the test suite + a smoke scrape after.

- [ ] **Step 1: Add module-level constant and thread-safe rate-limit state**

Near the top of `shein_scraper.py`, next to the other constants, add:

```python
MAX_PARALLEL_TABS      = 3        # 同时打开多少个 CDP tab 抓取（保守）
```

Below the `class RateLimitError(Exception)` block, add:

```python
_rate_limit_lock  = threading.Lock()
_rate_limit_state = {"consecutive_fails": 0, "tripped": False}


def _record_result_for_rate_limit(rec_status: str) -> bool:
    """Update the shared rate-limit counter. Returns True iff limiter has
    tripped (caller should abandon remaining URLs)."""
    with _rate_limit_lock:
        if _rate_limit_state["tripped"]:
            return True
        if rec_status != "OK":
            _rate_limit_state["consecutive_fails"] += 1
            if _rate_limit_state["consecutive_fails"] >= RATE_LIMIT_CONSECUTIVE:
                _rate_limit_state["tripped"] = True
                return True
        else:
            _rate_limit_state["consecutive_fails"] = 0
        return False


def _reset_rate_limit_state() -> None:
    with _rate_limit_lock:
        _rate_limit_state["consecutive_fails"] = 0
        _rate_limit_state["tripped"] = False
```

- [ ] **Step 2: Extract per-URL body into `_scrape_one_url`**

Cut the body inside the existing `for i, url in enumerate(urls, 1):` loop (from the top-of-loop `print(f"[{i}/{len(urls)}]…")` line through the `records.append(rec)` at the end of each iteration — excluding the rate-limit check and `_inter_url_pause` call, which stay in the outer driver) into a new function. Insert it above `scrape_shein`:

```python
def _scrape_one_url(
    url: str, index: int, total: int, seq_num: int,
    override_price, override_shipping, variant_filter_decl: str,
) -> dict:
    """Scrape one product URL in a fresh CDP tab. Returns the record dict.
    Thread-safe: each call opens its own tab and uses its own WebSocket."""
    print(f"[{index}/{total}] {url[:80]}...")
    rec = {"url": url, "status": "OK", "seq_num": seq_num}
    tab_id = None
    _url_start_time = time.monotonic()
    try:
        print("  navigating...")
        ws_url, tab_id = _navigate_and_wait(CDP_PORT, url)
        base_dir = Path.cwd()

        if not _check_and_handle_block(CDP_PORT, tab_id, url, base_dir):
            rec["status"] = "BLOCKED"
            print("  [跳过] 页面被拦截，无法提取")
            return rec
        _url_start_time = time.monotonic()

        _OOPS_DETECT_JS = """
            (function() {
                if (document.body && (
                    document.body.innerText.includes('Oops') ||
                    document.querySelector('.page-not-found, .error-page, [class*="not-found"]')
                )) return 'OOPS';
                if (document.title && document.title.includes('[goods_name]'))
                    return 'NO_DATA';
                return 'OK';
            })()
        """
        try:
            _page_check = _run_js(ws_url, _OOPS_DETECT_JS)
            if _page_check == "OOPS":
                _backoff = random.uniform(OOPS_RETRY_BACKOFF_MIN, OOPS_RETRY_BACKOFF_MAX)
                print(f"  [OOPS] 商品页显示 Oops — 退避 {_backoff:.0f}s 后重试...")
                time.sleep(_backoff)
                try:
                    ws_url = _reload_tab_and_wait(CDP_PORT, tab_id)
                    _page_check = _run_js(ws_url, _OOPS_DETECT_JS)
                except Exception:
                    pass
                if _page_check == "OOPS":
                    rec["status"] = "DELISTED"
                    print("  [跳过] 重试后仍 Oops，判定真下架")
                    return rec
                print(f"  [OOPS] 重试成功 (state={_page_check})，继续提取")
                _url_start_time = time.monotonic()

            if _page_check == "NO_DATA":
                rec["status"] = "NO_DATA"
                print("  [跳过] 页面数据未加载 ([goods_name])")
                return rec
        except Exception:
            pass

        try:
            _run_js(ws_url, _JS_SCROLL_GALLERY)
            time.sleep(1.0)
            ws_url = _ws_url_for_id(CDP_PORT, tab_id)
        except Exception:
            pass

        data = None
        for attempt in range(PAGE_LOAD_RETRIES):
            _elapsed = time.monotonic() - _url_start_time
            if _elapsed > EXTRACTION_TIMEOUT_SEC:
                print(f"  [超时] 页面处理已超过 {EXTRACTION_TIMEOUT_SEC}s，可能被隐形拦截")
                _recheck = None
                try:
                    ws_url = _ws_url_for_id(CDP_PORT, tab_id)
                    _recheck = _run_js(ws_url, _JS_DETECT_BLOCK)
                except Exception:
                    pass
                ss_path = str(_screenshots_dir(base_dir) /
                              f"_timeout_{time.strftime('%Y%m%d_%H%M%S')}.png")
                try:
                    ws_url = _ws_url_for_id(CDP_PORT, tab_id)
                    _take_screenshot(ws_url, ss_path)
                except Exception:
                    ss_path = None
                _block_info = ""
                if isinstance(_recheck, dict) and _recheck.get("blocked"):
                    _block_info = f"\n检测到拦截类型: {_recheck.get('type', 'unknown')}"
                alert_generic(
                    url,
                    f"页面处理超时（{_elapsed:.0f}s > {EXTRACTION_TIMEOUT_SEC}s），"
                    f"可能存在隐形验证码或页面加载异常。{_block_info}",
                    screenshot_path=ss_path,
                )
                rec["status"] = "TIMEOUT"
                data = None
                break

            if attempt == 0:
                _wait_for_shipping_text(CDP_PORT, tab_id, max_wait=8.0)
                ws_url = _ws_url_for_id(CDP_PORT, tab_id)
            print("  extracting...")
            data = _run_js(ws_url, _JS)
            if isinstance(data, dict) and not _shein_page_needs_retry(data):
                break
            if attempt < PAGE_LOAD_RETRIES - 1:
                print(f"  transient/error page (try {attempt + 1}/{PAGE_LOAD_RETRIES}), reloading...")
                time.sleep(RELOAD_PAUSE_SEC)
                ws_url = _reload_tab_and_wait(CDP_PORT, tab_id)

        if rec.get("status") == "TIMEOUT":
            print("  [跳过] 页面超时")
            return rec

        if not isinstance(data, dict):
            raise ValueError("JS returned unexpected type — page may not have loaded")

        data["variations"] = _merge_main_sale_attr_colors(
            data.get("variations") or {}, data.get("main_sale_attrs") or []
        )

        if variant_filter_decl:
            _kept, _filt_vars, _unknown = _filter_variants_by_declaration(
                variant_filter_decl,
                data.get("sku_prices") or [], data.get("variations") or {},
            )
            if _unknown:
                print(f"  [变体过滤] 声明中未匹配到: {_unknown}")
            if not _kept and (data.get("sku_prices") or []):
                rec["status"] = f"ERROR: variant_filter_no_match ({variant_filter_decl!r})"
                return rec
            data["sku_prices"] = _kept
            if _filt_vars:
                data["variations"] = _filt_vars

        web_price = data.get("price")
        if override_price is not None:
            op = float(override_price)
            data["price"] = op
            for _sp in (data.get("sku_prices") or []):
                _sp["sale_price"] = op
            print(f"  [price override] C列=${op:.2f} "
                  f"(网页=${(web_price or 0):.2f}, "
                  f"sku_prices×{len(data.get('sku_prices') or [])} 同步)")

        if override_shipping is not None:
            shipping = float(override_shipping)
            ship_note = f"template D=${shipping:.2f} (override)"
        else:
            shipping = _calc_shipping(data)
            ship_note = None
        price = data.get("price") or 0.0
        ebay  = _ebay_listing_price(price, shipping)
        thresh = data.get("free_threshold")
        if ship_note is None:
            if data.get("unconditional_free"):
                ship_note = "unconditional FREE shipping"
            elif thresh is not None:
                ship_note = (
                    f"threshold AU${thresh:.2f} — price AU${price:.2f} "
                    + ("≥ threshold → FREE" if price >= thresh
                       else f"< threshold → AU${DEFAULT_SHIPPING_FEE}")
                )
            elif not (data.get("shipping_raw") or "").strip():
                ship_note = "no shipping info on page → assumed FREE"
            else:
                ship_note = f"shipping text present but no threshold → AU${DEFAULT_SHIPPING_FEE}"

        sku_prices = data.get("sku_prices") or []
        rec.update({
            "sku":              data.get("goods_sn") or data.get("goods_id", ""),
            "price":            price,
            "web_price":        web_price,
            "web_price_display": _format_price_range(sku_prices) or (
                                  f"${float(web_price):.2f}"
                                  if web_price is not None else ""),
            "stock_summary":    _format_stock_summary(sku_prices),
            "shipping":         shipping,
            "shipping_raw":     data.get("shipping_raw") or "",
            "free_threshold":   data.get("free_threshold"),
            "unconditional_free": data.get("unconditional_free", False),
            "website":          "au.shein.com",
            "store_name":       data.get("store_name", ""),
            "original_title":   data.get("title", ""),
            "title":            data.get("title", ""),
            "variations":       data.get("variations", {}),
            "ebay_price":       ebay,
            "media":            data.get("media", {}),
            "goods_imgs":       data.get("goods_imgs") or [],
            "sku_prices":       sku_prices,
            "main_sale_attrs":  data.get("main_sale_attrs") or [],
            "seq_num":          seq_num,
        })

        try:
            vars_ = rec.get("variations") or {}
            color_key = next((k for k in vars_.keys() if "color" in (k or "").lower()), None)
            if color_key:
                cv = vars_.get(color_key) or []
                if isinstance(cv, list) and len(cv) == 1:
                    color_val = _clean_ws(str(cv[0]))
                    if color_val and color_val.lower() not in (rec["title"] or "").lower():
                        rec["title"] = _clean_ws(f"{rec['title']} - {color_val}")
        except Exception:
            pass
        if not rec["title"]:
            rec["status"] = "PARSE_ERROR"

        if sku_prices:
            for sp in sku_prices:
                for ak, av in (sp.get("attrs") or {}).items():
                    if not ak or not av:
                        continue
                    exists = any(vk.lower() == ak.lower() for vk in rec["variations"])
                    if not exists:
                        all_vals, seen_v = [], set()
                        for sp2 in sku_prices:
                            v2 = (sp2.get("attrs") or {}).get(ak, "")
                            if v2 and v2 not in seen_v:
                                seen_v.add(v2)
                                all_vals.append(v2)
                        if all_vals:
                            rec["variations"][ak] = all_vals

        _orig_t = rec.get("original_title", "")
        _ai_title = _make_ebay_title_ai(_orig_t)
        if _ai_title:
            rec["ebay_title"] = _ai_title
            print(f"  eBay title : {_ai_title} (AI)")
        else:
            rec["ebay_title"] = _make_ebay_title(_orig_t, rec.get("variations", {}))
            print(f"  eBay title : {rec['ebay_title']} (fallback)")

        seq_display = rec["seq_num"]
        goods_imgs_count = len(rec.get("goods_imgs") or [])
        print(f"  seq        : {seq_display}")
        print(f"  title      : {rec['title'][:65]}")
        print(f"  sku        : {rec['sku']}")
        print(f"  price      : ${price:.2f}")
        print(f"  shipping   : ${shipping:.2f}  ({ship_note})")
        print(f"  store      : {rec['store_name']}")
        print(f"  eBay price : ${ebay:.2f}")
        print(f"  variations : {rec['variations']}")
        print(f"  sku_prices : {len(sku_prices)} 个变体")
        print(f"  goods_imgs : {goods_imgs_count} 张（JSON解析）")

        media_folder = _download_media(rec, base_dir, seq_num=seq_display)
        _write_ebay_listing_txt(rec, media_folder)
        first_img = _first_product_image_path(media_folder)
        rec["first_image_path"] = str(first_img) if first_img else ""
        if media_folder:
            n_webp = len(list(Path(media_folder).glob("img_*.webp")))
            n_img  = len(list(Path(media_folder).glob("img_*.*")))
            n_vid  = len(list(Path(media_folder).glob("video_*.mp4")))
            print(f"  media      : {n_img} image(s) ({n_webp} webp), {n_vid} video(s)")

    except Exception as e:
        rec["status"] = f"ERROR: {e}"
        print(f"  ERROR: {e}")
    finally:
        if tab_id is not None:
            _close_tab(CDP_PORT, tab_id)

    return rec
```

- [ ] **Step 3: Replace the main loop in `scrape_shein` with ThreadPoolExecutor batches**

In `scrape_shein`, after the `_ensure_chrome()` / session-warmup lines (the print statements + `records = []`), replace the entire existing `for i, url in enumerate(urls, 1):` block through the closing `_inter_url_pause(i, len(urls))` line with:

```python
    _reset_rate_limit_state()
    _rate_limited = False

    # Session warmup once before the pool.
    _ensure_shein_session(CDP_PORT)

    # state_tracker (optional module) — thread-safe as a coarse-grained progress ping.
    try:
        import state_tracker as _st
    except Exception:
        _st = None

    _st_lock = threading.Lock()
    _done_count = 0

    def _bump_progress(url_hint: str) -> None:
        nonlocal _done_count
        with _st_lock:
            _done_count += 1
            done = _done_count
        if _st is not None:
            try:
                _st.url_progress(done=done, total=len(urls), current_url=url_hint)
            except Exception:
                pass

    total = len(urls)
    batch_size = MAX_PARALLEL_TABS
    for batch_start in range(0, total, batch_size):
        batch_end = min(batch_start + batch_size, total)
        # Check rate-limit trip before submitting the batch.
        with _rate_limit_lock:
            if _rate_limit_state["tripped"]:
                _rate_limited = True
                break

        with ThreadPoolExecutor(max_workers=batch_size) as ex:
            futures = []
            for i in range(batch_start, batch_end):
                seq_num = seq_list[i] if seq_list else start_seq + i
                op = price_list[i] if price_list else None
                os_ = shipping_list[i] if shipping_list else None
                vf = variant_filter_list[i] if variant_filter_list else ""
                futures.append(ex.submit(
                    _scrape_one_url,
                    urls[i], i + 1, total, seq_num, op, os_, vf,
                ))
            for fut in as_completed(futures):
                try:
                    rec = fut.result()
                except Exception as e:
                    rec = {"status": f"ERROR: {e}",
                           "url": "<unknown>", "seq_num": None}
                records.append(rec)
                _bump_progress(rec.get("url", ""))
                if _record_result_for_rate_limit(rec.get("status") or "OK"):
                    break

        # If limiter tripped mid-batch, mark all not-yet-scraped as RATE_LIMITED.
        with _rate_limit_lock:
            if _rate_limit_state["tripped"]:
                already_seen = {r.get("seq_num") for r in records}
                for j in range(batch_end, total):
                    seq_num = seq_list[j] if seq_list else start_seq + j
                    if seq_num in already_seen:
                        continue
                    records.append({
                        "url": urls[j], "status": "RATE_LIMITED",
                        "seq_num": seq_num,
                    })
                print(f"\n  [限流] 达到 {RATE_LIMIT_CONSECUTIVE} 次连续失败阈值，停止后续 URL")
                _rate_limited = True
                break

        # Pause 3s between batches (not after the final batch).
        if batch_end < total:
            _inter_url_pause(batch_end, total)
```

Keep the surrounding `try: … finally:` and the `_save_excel` / `RateLimitError` blocks intact.

- [ ] **Step 4: Sort records by seq_num before expanding/saving**

Because concurrent completion means `records` is out of order, add just before `expanded = _expand_records(records)`:

```python
    records.sort(key=lambda r: (r.get("seq_num") or 0))
```

- [ ] **Step 5: Verify module imports and existing tests still pass**

```powershell
python -c "import shein_scraper; import inspect; s=inspect.signature(shein_scraper.scrape_shein); print(s)"
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
```

Expected: all four test files print `ALL PASS`; the signature output includes `shipping_list=None, variant_filter_list=None`.

- [ ] **Step 6: Commit**

```powershell
git add shein_scraper.py
git commit -m "feat(scraper): 3-tab parallelism via ThreadPoolExecutor + thread-safe rate limiter"
```

---

## Task 11: End-to-end manual smoke test + AI-key packaging verification

**Files:** none new — this task validates the whole plan against real behavior.

- [ ] **Step 1: Prepare a smoke template**

Copy `C:\Users\ak\Downloads\Template - Test.xlsx` to your `SHEIN_SUBMITTED_DIR` (or set `SHEIN_INPUT_FILENAME=Template - Test.xlsx` in `.env`). Fill 5 rows on the `ZR1` sheet:

- **Row 2 (single-variant product):** 编号=1, 链接=<known-good URL>, 原价=10.00, 运费=7.95, 变体=(empty).
- **Row 3 (multi-variant, filter all):** 编号=2, 链接=<multi-color product URL>, 原价=12.00, 运费=7.95, 变体=(empty).
- **Row 4 (multi-variant, filter subset):** 编号=3, 链接=<same or another multi-color URL>, 原价=12.00, 运费=7.95, 变体="Black, Red / M, L" (use values you know exist on the page).
- **Row 5 (delisted probe):** 编号=4, 链接=<known-delisted URL>, 原价=9.99, 运费=7.95, 变体=(empty).
- **Row 6 (variant filter mismatch):** 编号=5, 链接=<multi-variant URL>, 原价=12.00, 运费=7.95, 变体="Purple / XXL" (values that don't exist).

- [ ] **Step 2: Run source-mode scrape**

```powershell
python run_excel.py "<path-to-template>"
```

- [ ] **Step 3: Verify template rows F–L**

Open the template. For each row, expect:

- **Row 2 (single-variant):** F=today, G=Done, H=`$X.XX` (single price), I=raw h1, J=AI-generated (does **not** start with "NEW " unless the AI put it there), K=`10 × 2 + 7.95 = 27.95`, L=`<attr>: <n>` or `<attr>: 少货 <n>` / `<attr>: 缺货`.
- **Row 3 (multi, no filter):** F=today, G=Done, H=range like `$9.99–$14.99`, K=`12 × 2 + 7.95 = 31.95`, L=`Black-M: 12 / Black-L: 缺货 / Red-M: 少货 3 / …`.
- **Row 4 (multi, filter):** L contains **only** Black/Red × M/L combinations, not other colors.
- **Row 5 (delisted):** F=today, G=Delisted, all other cols empty.
- **Row 6 (filter mismatch):** F=today, G=Failed, all other cols empty.

Also verify:
- `OUTPUT_ROOT/ZR1/ZR1-1-5-<yyyymmdd>.xlsx` exists with the old **English-header** 14-column layout.
- Each successful `图片-<sku>/` folder has `eBay上架描述.txt` with the same eBay 标题 as the template's J column and the same eBay 价格 as K.
- Wall-clock for the 5 URLs is roughly 5/3 batches × per-URL time — noticeably faster than serial.

If any row shows the fallback title pattern (starts with `NEW ` and looks like a raw truncation), the AI path failed — check `.env` has `ANTHROPIC_API_KEY=sk-ant-…` (source mode uses `.env`; installer mode uses `key_store.py`).

- [ ] **Step 4: Verify installer bundling (item 1's real acceptance test)**

Create `.build_key.txt` with your Anthropic key on one line, then:

```powershell
build.bat exe
```

Expected: 3 steps run in order (`make_key_store.py`, PyInstaller, skip Inno Setup on `exe` arg), no errors, `dist\SheinExtractAU.exe` produced.

Verify `key_store` is inside the exe:

```powershell
dist\SheinExtractAU.exe --version 2>&1 | Select-Object -First 5
# then in a fresh terminal, use the exe to scrape one row of the smoke template
# and confirm the J column is AI-generated, not the fallback pattern.
```

Clean up:

```powershell
Remove-Item .build_key.txt, key_store.py -Force -ErrorAction SilentlyContinue
```

- [ ] **Step 5: Commit the plan checklist completion**

No code change here — this task is verification only. If anything failed, open a follow-up before merging.

---

## Rollout

- Merge `feat/template-refactor` to `main` after Task 11 passes.
- Bump version (`version.py`), run `release.bat X.Y.Z` (existing owner workflow), and distribute the new installer to employees.
- Announce: old English-header input files no longer processed; use the new Chinese-header template only.
