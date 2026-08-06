# eBay Price Check Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Post-scrape helper `python ebay_price_check.py` that, for each `G=Done AND N=empty` row, searches `ebay.com.au` (3-tab parallel CDP), extracts top-2 Best Match new/BIN/AU listings, and writes delivered price + cleaned URL to columns N–R for human 2审 of L-column pricing.

**Architecture:**
- Two new files: `ebay_scraper.py` (pure/library — URL cleaner, price/postage parsers, `_extract_from_page`, `search_ebay_au`, `_ensure_ebay_session`) and `ebay_price_check.py` (CLI + driver — reader/writer, `ThreadPoolExecutor(max_workers=3)` batch loop, session warmup, write-back).
- Reuses `shein_scraper.py` verbatim for CDP plumbing (`_ensure_chrome`, `_new_tab`, `_close_tab`, `_ws_url_for_id`, `_run_js`, `_cdp_once`), pacing (`_inter_url_pause`, `INTER_URL_DELAY_SEC`), parallelism knobs (`MAX_PARALLEL_TABS`, `_rate_limit_lock`, `_record_result_for_rate_limit`, `_reset_rate_limit_state`), and screenshot/notification helpers (`_take_screenshot`, `_screenshots_dir`, `alert_captcha`).
- One targeted modification to `shein_scraper.py::_check_and_handle_block` to add eBay-specific captcha/sign-in DOM selectors — kept in the shared function to preserve single-source detection.

**Tech Stack:** Python 3.10+, openpyxl, requests, websocket-client, Chrome DevTools Protocol. Tests are plain-`assert` scripts (project convention — see `test_variant_merge.py`), runnable via `python test_<name>.py`.

**Working branch:** `feat/ebay-price-check` (already checked out at commit `16237aa`).

**Reference spec:** `docs/superpowers/specs/2026-08-02-ebay-price-check-design.md`.

---

## File Structure

**New files:**
- `ebay_scraper.py` — library module:
  - `EbayHit` dataclass (`title`, `sticker_price`, `postage`, `delivered_price`, `url`, `image_url`)
  - `_clean_ebay_url(u: str) -> str`
  - `_parse_price(s: str) -> float | None`
  - `_parse_postage(s: str) -> float | None`
  - `_JS_EBAY_SEARCH_EXTRACT` — extraction JS payload
  - `_extract_from_page(raw_hits: list[dict]) -> list[EbayHit]`
  - `_ensure_ebay_session(port: int) -> None` (session warmup, `_EBAY_SESSION_WARMED` gate)
  - `search_ebay_au(query: str, port: int) -> list[EbayHit]` (public API)
- `ebay_price_check.py` — CLI + driver:
  - Column constants `COL_EBAY_SEARCH_DATE=14`, `COL_LOW_PRICE=15`, `COL_LOW_URL=16`, `COL_HIGH_PRICE=17`, `COL_HIGH_URL=18`
  - `_read_ebay_pending_rows(ws) -> list[dict]`
  - `_write_ebay_result(ws, row, search_date, low_price=None, low_url=None, high_price=None, high_url=None) -> None`
  - `_check_one_row(row: dict, port: int) -> dict` (worker)
  - `process_excel(xlsx_path: Path) -> None`
  - `main()`
- `test_ebay_url_clean.py` — URL cleaner tests
- `test_ebay_parse.py` — price/postage parser tests
- `test_ebay_extract.py` — `_extract_from_page` tests using synthetic raw-hit dicts
- `test_ebay_io.py` — reader/writer tests using openpyxl in-memory workbook

**Modified files:**
- `shein_scraper.py::_check_and_handle_block` — extend selectors to catch eBay captcha (Arkose iframe `iframe[src*="arkoselabs"]`, `iframe[src*="captcha"]` already there) and eBay sign-in overlays (`.signin-container`, `[class*="signin"][class*="overlay"]`). Existing generic patterns already cover most cases; small additive change.

**No changes to:** `run_excel.py`, `config.py`, `.env` (script picks up `SHEIN_INPUT_FILENAME` and `SHEIN_SUBMITTED_DIR` transparently), tests for shein/template.

---

## Task 1: URL cleaner — TDD

**Files:**
- Create: `test_ebay_url_clean.py`
- Create (skeleton): `ebay_scraper.py`

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_ebay_url_clean.py`:

```python
"""Tests for _clean_ebay_url — strips tracking / query params from
eBay listing URLs, keeps a clean sharable form."""
from ebay_scraper import _clean_ebay_url


def test_strip_campaign_param():
    u = "https://www.ebay.com.au/itm/123456789?campid=5338&customid=xyz"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/123456789"


def test_strip_hash_param():
    u = "https://www.ebay.com.au/itm/12345?hash=item1abc:g:XYZAA"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/12345"


def test_strip_trkparms_and_all_others():
    u = ("https://www.ebay.com.au/itm/12345"
         "?_trkparms=abc%3Dxyz&mkevt=1&mkcid=1&mkrid=705-53470-19255-0"
         "&campid=5338&customid=default")
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/12345"


def test_preserve_slug_variant():
    # eBay sometimes uses /itm/<slug>/<id> — keep both.
    u = "https://www.ebay.com.au/itm/Cute-Ceramic-Mug/22334455?campid=5338"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/Cute-Ceramic-Mug/22334455"


def test_already_clean_passthrough():
    u = "https://www.ebay.com.au/itm/98765"
    assert _clean_ebay_url(u) == u


def test_non_ebay_passthrough():
    # Defensive: if extractor ever picks up a non-eBay href, don't mangle it.
    u = "https://example.com/some?path=1&other=2"
    assert _clean_ebay_url(u) == u


if __name__ == "__main__":
    test_strip_campaign_param()
    test_strip_hash_param()
    test_strip_trkparms_and_all_others()
    test_preserve_slug_variant()
    test_already_clean_passthrough()
    test_non_ebay_passthrough()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ModuleNotFoundError**

```powershell
python test_ebay_url_clean.py
```

Expected: `ModuleNotFoundError: No module named 'ebay_scraper'`.

- [ ] **Step 3: Create `ebay_scraper.py` with the function**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\ebay_scraper.py`:

```python
"""eBay AU search scraper (post-Shein price-check helper).

Public API:
    search_ebay_au(query: str, port: int) -> list[EbayHit]
    _ensure_ebay_session(port: int) -> None
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from urllib.parse import urlparse, urlunparse


def _clean_ebay_url(u: str) -> str:
    """Strip the entire query string from an eBay listing URL.
    Non-eBay URLs pass through unchanged (defensive)."""
    if not isinstance(u, str) or not u:
        return u
    try:
        p = urlparse(u)
    except Exception:
        return u
    host = (p.netloc or "").lower()
    if "ebay." not in host:
        return u
    # Drop query + fragment; keep scheme, host, path.
    return urlunparse((p.scheme, p.netloc, p.path, "", "", ""))
```

- [ ] **Step 4: Run — expect PASS**

```powershell
python test_ebay_url_clean.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_ebay_url_clean.py ebay_scraper.py
git commit -m "feat(ebay): _clean_ebay_url strips tracking params from listing URLs"
```

---

## Task 2: Price parser — TDD

**Files:**
- Create: `test_ebay_parse.py`
- Modify: `ebay_scraper.py`

- [ ] **Step 1: Write the failing tests**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_ebay_parse.py`:

```python
"""Tests for _parse_price (sticker) and _parse_postage (shipping) —
raw-string → float, robust to eBay AU price formats."""
from ebay_scraper import _parse_price, _parse_postage


# ── _parse_price ────────────────────────────────────────────────────────────

def test_price_au_prefix():
    assert _parse_price("AU $12.99") == 12.99


def test_price_bare_dollar():
    assert _parse_price("$12.99") == 12.99


def test_price_thousand_separator():
    assert _parse_price("AU $1,234.99") == 1234.99


def test_price_integer_no_decimals():
    assert _parse_price("AU $50") == 50.0


def test_price_range_uses_low_end():
    # 'AU $12.99 to AU $18.99' — take the lowest.
    assert _parse_price("AU $12.99 to AU $18.99") == 12.99


def test_price_whitespace_tolerant():
    assert _parse_price("  AU $ 12.99  ") == 12.99


def test_price_empty_returns_none():
    assert _parse_price("") is None
    assert _parse_price(None) is None


def test_price_gibberish_returns_none():
    assert _parse_price("SEE PRICE IN CART") is None


# ── _parse_postage ──────────────────────────────────────────────────────────

def test_postage_free():
    assert _parse_postage("Free postage") == 0.0


def test_postage_free_case_insensitive():
    assert _parse_postage("FREE POSTAGE") == 0.0
    assert _parse_postage("free shipping") == 0.0


def test_postage_plus_prefix():
    assert _parse_postage("+ AU $6.95 postage") == 6.95


def test_postage_no_plus():
    assert _parse_postage("AU $6.95 postage") == 6.95


def test_postage_empty_returns_none():
    assert _parse_postage("") is None
    assert _parse_postage(None) is None


def test_postage_unparseable_returns_none():
    assert _parse_postage("Contact seller for postage") is None


if __name__ == "__main__":
    test_price_au_prefix()
    test_price_bare_dollar()
    test_price_thousand_separator()
    test_price_integer_no_decimals()
    test_price_range_uses_low_end()
    test_price_whitespace_tolerant()
    test_price_empty_returns_none()
    test_price_gibberish_returns_none()
    test_postage_free()
    test_postage_free_case_insensitive()
    test_postage_plus_prefix()
    test_postage_no_plus()
    test_postage_empty_returns_none()
    test_postage_unparseable_returns_none()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ImportError**

```powershell
python test_ebay_parse.py
```

Expected: `ImportError: cannot import name '_parse_price' from 'ebay_scraper'`.

- [ ] **Step 3: Add the two parsers**

Append to `ebay_scraper.py` (after `_clean_ebay_url`):

```python
# Matches "12.99", "1,234.99", "50" — captures the numeric literal.
_PRICE_NUM_RE = re.compile(r"(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)")


def _parse_price(s) -> "float | None":
    """Parse an eBay AU sticker price string into a float.
    Handles 'AU $12.99', '$12.99', 'AU $1,234.99', integer prices, ranges
    ('AU $X to AU $Y' → take X), whitespace. Returns None for None / empty
    / no numeric substring."""
    if s is None:
        return None
    if not isinstance(s, str):
        return None
    t = s.strip()
    if not t:
        return None
    m = _PRICE_NUM_RE.search(t)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except (ValueError, TypeError):
        return None


def _parse_postage(s) -> "float | None":
    """Parse eBay postage cell text. 'Free postage' / 'Free shipping' → 0.0.
    '+ AU $6.95 postage' / 'AU $6.95 postage' → 6.95. Returns None if
    unparseable or empty."""
    if s is None or not isinstance(s, str):
        return None
    t = s.strip()
    if not t:
        return None
    if re.search(r"\bfree\s+(postage|shipping|delivery)\b", t, re.I):
        return 0.0
    return _parse_price(t)  # reuse the price extractor for the numeric part
```

- [ ] **Step 4: Run — expect PASS**

```powershell
python test_ebay_parse.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_ebay_parse.py ebay_scraper.py
git commit -m "feat(ebay): _parse_price + _parse_postage handle AU $ formats and Free shipping"
```

---

## Task 3: `_extract_from_page` — TDD (converts raw JS return to `EbayHit`s)

**Files:**
- Create: `test_ebay_extract.py`
- Modify: `ebay_scraper.py`

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_ebay_extract.py`:

```python
"""Tests for _extract_from_page — takes the JS return (list of dicts with
raw strings from the eBay search page) and yields a normalized list of
EbayHit instances, cleaning URLs and computing delivered_price."""
from ebay_scraper import _extract_from_page, EbayHit


def _raw(title="Foo Widget", price="AU $12.99", postage="+ AU $6.95 postage",
         url="https://www.ebay.com.au/itm/111?campid=5338&hash=x",
         img="https://i.ebayimg.com/thumbs/foo.jpg"):
    return {"title": title, "price": price, "postage": postage,
            "url": url, "image_url": img}


def test_two_full_hits():
    raw = [_raw(title="A"), _raw(title="B", price="AU $15.00", postage="Free postage",
                                 url="https://www.ebay.com.au/itm/222?campid=1")]
    hits = _extract_from_page(raw)
    assert len(hits) == 2
    assert hits[0].title == "A"
    assert hits[0].sticker_price == 12.99
    assert hits[0].postage == 6.95
    assert hits[0].delivered_price == 19.94
    assert hits[0].url == "https://www.ebay.com.au/itm/111"
    assert hits[1].sticker_price == 15.00
    assert hits[1].postage == 0.0
    assert hits[1].delivered_price == 15.00


def test_postage_unparseable_falls_back_to_sticker():
    raw = [_raw(postage="Contact seller for postage")]
    hits = _extract_from_page(raw)
    assert len(hits) == 1
    assert hits[0].postage is None
    assert hits[0].delivered_price == hits[0].sticker_price  # = 12.99


def test_price_unparseable_skips_that_hit():
    # No usable price → drop this hit (can't rank without a price).
    raw = [_raw(price="SEE PRICE IN CART"), _raw(title="Real Hit")]
    hits = _extract_from_page(raw)
    assert len(hits) == 1
    assert hits[0].title == "Real Hit"


def test_empty_input_returns_empty_list():
    assert _extract_from_page([]) == []


def test_at_most_two_returned_even_if_more_input():
    raw = [_raw(title=f"Hit{i}") for i in range(5)]
    hits = _extract_from_page(raw)
    assert len(hits) == 2
    assert hits[0].title == "Hit0"
    assert hits[1].title == "Hit1"


def test_url_is_cleaned():
    raw = [_raw(url="https://www.ebay.com.au/itm/Nice-Slug/333?campid=5338&_trkparms=x")]
    hits = _extract_from_page(raw)
    assert hits[0].url == "https://www.ebay.com.au/itm/Nice-Slug/333"


if __name__ == "__main__":
    test_two_full_hits()
    test_postage_unparseable_falls_back_to_sticker()
    test_price_unparseable_skips_that_hit()
    test_empty_input_returns_empty_list()
    test_at_most_two_returned_even_if_more_input()
    test_url_is_cleaned()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ImportError**

```powershell
python test_ebay_extract.py
```

Expected: `ImportError: cannot import name '_extract_from_page' from 'ebay_scraper'`.

- [ ] **Step 3: Add `EbayHit` + `_extract_from_page`**

Append to `ebay_scraper.py` (after `_parse_postage`):

```python
@dataclass
class EbayHit:
    title: str
    sticker_price: float
    postage: "float | None"  # None → couldn't parse; delivered falls back to sticker
    delivered_price: float
    url: str
    image_url: str


def _extract_from_page(raw_hits: list) -> "list[EbayHit]":
    """Convert the JS-extracted raw hits into normalized EbayHit list.
    Drops entries without a parseable sticker price. Returns at most 2
    (Best Match rank 1 and rank 2). URL cleaned, postage-fail falls back
    to sticker for delivered."""
    out: list[EbayHit] = []
    for raw in raw_hits or []:
        if len(out) >= 2:
            break
        sticker = _parse_price(str(raw.get("price") or ""))
        if sticker is None:
            continue  # can't rank without a price
        postage = _parse_postage(str(raw.get("postage") or ""))
        delivered = sticker + (postage if postage is not None else 0.0)
        # If postage was truly missing (None), delivered = sticker only.
        if postage is None:
            delivered = sticker
        out.append(EbayHit(
            title=str(raw.get("title") or "").strip(),
            sticker_price=sticker,
            postage=postage,
            delivered_price=round(delivered, 2),
            url=_clean_ebay_url(str(raw.get("url") or "")),
            image_url=str(raw.get("image_url") or "").strip(),
        ))
    return out
```

- [ ] **Step 4: Run — expect PASS**

```powershell
python test_ebay_extract.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_ebay_extract.py ebay_scraper.py
git commit -m "feat(ebay): EbayHit dataclass + _extract_from_page normalizes raw JS return"
```

---

## Task 4: JS extraction payload + `search_ebay_au` (integration wrapper)

**Files:**
- Modify: `ebay_scraper.py`

- [ ] **Step 1: Add the JS extraction constant**

Append to `ebay_scraper.py`:

```python
# ── JavaScript: extract top search results from ebay.com.au SRP ─────────────
_JS_EBAY_SEARCH_EXTRACT = r"""
(function() {
    // Card selectors: eBay AU uses .s-item within #srp-river-results as of 2026.
    // Selector chain is defensive — fall back to older classes if the primary
    // isn't found.
    const cards = document.querySelectorAll(
        '#srp-river-results li.s-item, ul.srp-results li.s-item, li.s-item'
    );
    const out = [];
    for (const c of cards) {
        // Skip the "Shop on eBay" placeholder card that eBay sometimes injects.
        const t = (c.querySelector('.s-item__title, [class*="title"]')?.innerText || '').trim();
        if (!t || /^shop on ebay$/i.test(t)) continue;

        const priceEl = c.querySelector('.s-item__price, [class*="price"]');
        const price = (priceEl?.innerText || '').trim();

        // Postage lives in a nearby element — labels vary. Try several.
        const postEl =
            c.querySelector('.s-item__shipping, .s-item__logisticsCost') ||
            [...c.querySelectorAll('span,div')].find(e =>
                /postage|shipping|delivery/i.test(e.innerText || ''));
        const postage = (postEl?.innerText || '').trim();

        // The listing link is the .s-item__link anchor.
        const linkEl = c.querySelector('a.s-item__link, a[href*="/itm/"]');
        const url = linkEl?.href || '';

        // First thumbnail image.
        const imgEl = c.querySelector('img.s-item__image-img, img[src*="ebayimg"]');
        const image_url = imgEl?.src || '';

        if (t && price && url) {
            out.push({title: t, price: price, postage: postage, url: url, image_url: image_url});
        }
        if (out.length >= 5) break;  // grab a small pool; extractor keeps top 2
    }
    return out;
})()
"""
```

- [ ] **Step 2: Add `search_ebay_au` and `_ensure_ebay_session`**

Also append to `ebay_scraper.py`:

```python
# Import CDP helpers from shein_scraper. These are stable, module-level.
from shein_scraper import (
    CDP_PORT,
    _cdp_once,
    _new_tab,
    _close_tab,
    _ws_url_for_id,
    _run_js,
)


_EBAY_SESSION_WARMED = False


def _ensure_ebay_session(port: int = CDP_PORT) -> None:
    """Warm the browser session by loading ebay.com.au once. Idempotent —
    later calls in the same process are no-ops. Runs before the pool starts."""
    global _EBAY_SESSION_WARMED
    if _EBAY_SESSION_WARMED:
        return
    import requests, time
    tabs = requests.get(f"http://localhost:{port}/json", timeout=5).json()
    pages = [t for t in tabs if t.get("type") == "page"]
    ebay_pages = [t for t in pages if "ebay." in (t.get("url") or "")]
    if ebay_pages:
        _EBAY_SESSION_WARMED = True
        return
    print("  [导航] 预热：先访问 ebay.com.au 首页建立 session...")
    tab_id, ws_url = _new_tab(port, url="https://www.ebay.com.au/")
    time.sleep(4)
    _close_tab(port, tab_id)
    _EBAY_SESSION_WARMED = True


def search_ebay_au(query: str, port: int = CDP_PORT) -> "list[EbayHit]":
    """Open a fresh tab, run an eBay AU search with Best Match / New / BIN /
    AU-located filters, extract the top-2 listings. Returns [] on zero
    results. Raises RuntimeError on captcha (caller handles retry)."""
    from urllib.parse import quote
    from shein_scraper import _JS_DETECT_BLOCK  # reuse existing block detector

    q = quote((query or "").strip(), safe="")
    if not q:
        return []
    url = (
        "https://www.ebay.com.au/sch/i.html"
        f"?_nkw={q}"
        "&LH_BIN=1"
        "&LH_ItemCondition=1000"
        "&LH_PrefLoc=1"
        "&_sop=12"
    )

    tab_id, ws_url = _new_tab(port, url=url)
    if not tab_id:
        raise RuntimeError("Failed to open eBay search tab")
    try:
        import time
        # Wait for results grid — poll a few seconds. 2s min, up to 10s.
        deadline = time.monotonic() + 10.0
        time.sleep(2)
        raw_hits = []
        while time.monotonic() < deadline:
            ws_url = _ws_url_for_id(port, tab_id)
            # Captcha check
            try:
                blk = _run_js(ws_url, _JS_DETECT_BLOCK)
                if isinstance(blk, dict) and blk.get("blocked"):
                    raise RuntimeError(f"eBay blocked: {blk.get('type')}")
            except RuntimeError:
                raise
            except Exception:
                pass
            try:
                raw_hits = _run_js(ws_url, _JS_EBAY_SEARCH_EXTRACT) or []
                if raw_hits:
                    break
            except Exception:
                pass
            time.sleep(1.0)
        return _extract_from_page(raw_hits)
    finally:
        _close_tab(port, tab_id)
```

- [ ] **Step 3: Quick smoke — module imports without error**

```powershell
python -c "import ebay_scraper; print('ok'); print('EbayHit:', ebay_scraper.EbayHit); print('search_ebay_au:', ebay_scraper.search_ebay_au)"
```

Expected: `ok`, then two lines confirming class + function present.

- [ ] **Step 4: Re-run existing tests**

```powershell
python test_ebay_url_clean.py
python test_ebay_parse.py
python test_ebay_extract.py
```

Expected: three `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add ebay_scraper.py
git commit -m "feat(ebay): search_ebay_au + _ensure_ebay_session + extraction JS"
```

---

## Task 5: Extend `_check_and_handle_block` with eBay selectors

**Files:**
- Modify: `shein_scraper.py::_JS_DETECT_BLOCK` and `_JS_DISMISS_SIGNIN`

- [ ] **Step 1: Read the current detector JS**

Find `_JS_DETECT_BLOCK` in `shein_scraper.py` (currently around line 175). It checks for Shein sign-in modals and generic captcha iframes. The generic captcha checks already cover Arkose (`iframe[src*="captcha"]`), so eBay's Arkose challenges are picked up. The gap is eBay's sign-in overlay and the "Please verify you're a human" page, which uses different selectors.

- [ ] **Step 2: Extend the signin-modal selector list**

Locate this line in `_JS_DETECT_BLOCK`:

```javascript
    var modal = document.querySelector(
        '.sui-dialog__wrapper, [class*="modal"][class*="login"], [class*="overlay"][class*="login"]'
    );
```

Add eBay-specific selectors to the same querySelector call:

```javascript
    var modal = document.querySelector(
        '.sui-dialog__wrapper, [class*="modal"][class*="login"], [class*="overlay"][class*="login"], ' +
        '.signin-container, [class*="signin"][class*="overlay"], [id*="signin"][class*="modal"]'
    );
```

- [ ] **Step 3: Extend the dismiss button selectors**

Locate `_JS_DISMISS_SIGNIN` in `shein_scraper.py` (currently around line 232). Extend its `closeSelectors` array to include eBay patterns. Change:

```javascript
    var closeSelectors = [
        '.sui-dialog__headerbtn', '.sui-icon-common__close',
        '[class*="modal"] [class*="close"]', '[class*="dialog"] [class*="close"]',
        '[class*="popup"] [class*="close"]', '[aria-label="Close"]',
        '[class*="login"] [class*="close"]', '.she-close',
        'button[class*="close"]', '.icon-close',
    ];
```

To:

```javascript
    var closeSelectors = [
        '.sui-dialog__headerbtn', '.sui-icon-common__close',
        '[class*="modal"] [class*="close"]', '[class*="dialog"] [class*="close"]',
        '[class*="popup"] [class*="close"]', '[aria-label="Close"]',
        '[class*="login"] [class*="close"]', '.she-close',
        'button[class*="close"]', '.icon-close',
        // eBay signin/promo overlays
        '[class*="signin"] [class*="close"]', '.mp-close', '.dlg-close',
        'button[aria-label*="close" i]', '[data-testid*="close"]',
    ];
```

- [ ] **Step 4: Run existing Shein tests — none should break**

```powershell
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
python test_ebay_url_clean.py
python test_ebay_parse.py
python test_ebay_extract.py
```

Expected: all seven print `ALL PASS`. (The Shein tests don't exercise DOM detection JS, so they'll pass unaffected; the added selectors are additive.)

- [ ] **Step 5: Commit**

```powershell
git add shein_scraper.py
git commit -m "feat(scraper): extend _check_and_handle_block DOM selectors for eBay signin overlays"
```

---

## Task 6: Row reader + writer for N–R columns — TDD

**Files:**
- Create: `test_ebay_io.py`
- Create: `ebay_price_check.py`

- [ ] **Step 1: Write the failing test**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\test_ebay_io.py`:

```python
"""Tests for ebay_price_check row reader (_read_ebay_pending_rows) and
writer (_write_ebay_result) against the template schema."""
from openpyxl import Workbook

from ebay_price_check import _read_ebay_pending_rows, _write_ebay_result


CN_HEADERS = ["编号", "链接", "原价", "运费", "变体",
              "日期", "状态", "图片", "希音价格", "希音标题",
              "eBay标题", "eBay价格", "库存",
              "eBay搜索日期", "eBay同类低价", "低价链接",
              "eBay同类高价", "高价链接",
              "Shein重跑日期", "更新价格", "更新库存"]
NCOLS = len(CN_HEADERS)  # 21


def _make_ws(rows: list):
    wb = Workbook()
    ws = wb.active
    ws.title = "ZR1"
    for ci, h in enumerate(CN_HEADERS, 1):
        ws.cell(1, ci).value = h
    for ri, r in enumerate(rows, 2):
        padded = list(r) + [None] * (NCOLS - len(r))
        for ci, v in enumerate(padded, 1):
            ws.cell(ri, ci).value = v
    return ws


# ── _read_ebay_pending_rows ─────────────────────────────────────────────────

def test_reads_done_rows_with_empty_search_date():
    # Row 2: Done + no search date → PICK.
    # Row 3: Failed → skip.
    # Row 4: Done + already searched → skip.
    # Row 5: Delisted → skip.
    ws = _make_ws([
        [1, "url1", 10, 7.95, "", "2026-08-02", "Done",
         None, 9.99, "shein-title", "ebay-title", 27.95, "M: 20", None],
        [2, "url2", 12, 0,    "", "2026-08-02", "Failed",
         None, None, None, None, None, None, None],
        [3, "url3", 15, 7.95, "", "2026-08-02", "Done",
         None, 14.99, "st3", "et3", 37.93, "one-size: 20", "2026-08-02"],  # N filled
        [4, "url4", 20, 7.95, "", "2026-08-02", "Delisted",
         None, None, None, None, None, None, None],
    ])
    pending = _read_ebay_pending_rows(ws)
    assert len(pending) == 1
    p = pending[0]
    assert p["row"] == 2
    assert p["seq"] == 1
    assert p["ebay_title"] == "ebay-title"
    assert p["shein_title"] == "shein-title"


def test_skips_row_with_empty_ebay_title():
    # G=Done, N=empty, but K (eBay 标题) is empty → skip (can't search).
    ws = _make_ws([
        [1, "url1", 10, 7.95, "", "2026-08-02", "Done",
         None, 9.99, "shein-title", None, 27.95, "M: 20", None],
    ])
    pending = _read_ebay_pending_rows(ws)
    assert pending == []


def test_case_sensitive_done_only():
    # 'done' (lowercase) → skip; only exact 'Done' picks.
    ws = _make_ws([
        [1, "url1", 10, 7.95, "", "2026-08-02", "done",
         None, 9.99, "st", "et", 27.95, "M: 20", None],
        [2, "url2", 10, 7.95, "", "2026-08-02", "Done",
         None, 9.99, "st", "et", 27.95, "M: 20", None],
    ])
    pending = _read_ebay_pending_rows(ws)
    assert [p["seq"] for p in pending] == [2]


def test_rejects_non_template_sheet():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "Seq"
    ws.cell(1, 2).value = "Website"  # English → not our schema
    pending = _read_ebay_pending_rows(ws)
    assert pending == []


# ── _write_ebay_result ──────────────────────────────────────────────────────

def test_write_two_results():
    ws = _make_ws([
        [1, "url1", 10, 7.95, "", "2026-08-02", "Done",
         None, 9.99, "st", "et", 27.95, "M: 20", None],
    ])
    _write_ebay_result(
        ws, row=2,
        search_date="2026-08-02",
        low_price=25.90, low_url="https://www.ebay.com.au/itm/111",
        high_price=32.50, high_url="https://www.ebay.com.au/itm/222",
    )
    assert ws.cell(2, 14).value == "2026-08-02"           # N
    assert ws.cell(2, 15).value == 25.90                   # O
    assert ws.cell(2, 16).value == "https://www.ebay.com.au/itm/111"  # P
    assert ws.cell(2, 17).value == 32.50                   # Q
    assert ws.cell(2, 18).value == "https://www.ebay.com.au/itm/222"  # R


def test_write_one_result_leaves_high_empty():
    ws = _make_ws([[1, "url1", 10, 7.95, "", "2026-08-02", "Done",
                    None, 9.99, "st", "et", 27.95, "M: 20", None]])
    _write_ebay_result(ws, row=2, search_date="2026-08-02",
                       low_price=25.90, low_url="https://www.ebay.com.au/itm/111")
    assert ws.cell(2, 14).value == "2026-08-02"
    assert ws.cell(2, 15).value == 25.90
    assert ws.cell(2, 16).value == "https://www.ebay.com.au/itm/111"
    assert ws.cell(2, 17).value is None
    assert ws.cell(2, 18).value is None


def test_write_zero_result_marks_no_match():
    ws = _make_ws([[1, "url1", 10, 7.95, "", "2026-08-02", "Done",
                    None, 9.99, "st", "et", 27.95, "M: 20", None]])
    _write_ebay_result(ws, row=2, search_date="2026-08-02", no_match=True)
    assert ws.cell(2, 14).value == "2026-08-02"
    assert ws.cell(2, 15).value == "no match"
    assert ws.cell(2, 16).value is None
    assert ws.cell(2, 17).value is None
    assert ws.cell(2, 18).value is None


if __name__ == "__main__":
    test_reads_done_rows_with_empty_search_date()
    test_skips_row_with_empty_ebay_title()
    test_case_sensitive_done_only()
    test_rejects_non_template_sheet()
    test_write_two_results()
    test_write_one_result_leaves_high_empty()
    test_write_zero_result_marks_no_match()
    print("ALL PASS")
```

- [ ] **Step 2: Run — expect ModuleNotFoundError**

```powershell
python test_ebay_io.py
```

Expected: `ModuleNotFoundError: No module named 'ebay_price_check'`.

- [ ] **Step 3: Create `ebay_price_check.py` skeleton with the two helpers**

Create `C:\Users\ak\Desktop\Claude\shein-extract-au\ebay_price_check.py`:

```python
"""Post-scrape helper: for each row where G=Done and N=empty, search
ebay.com.au via CDP browser (3-tab parallel), extract top-2 Best Match
listings, and write delivered price + cleaned URL to columns N-R.

Usage:
    python ebay_price_check.py                  # picks SUBMITTED_DIR/SHEIN_INPUT_FILENAME
    python ebay_price_check.py "path/to/x.xlsx" # explicit file

Columns written (see docs/superpowers/specs/2026-08-02-ebay-price-check-design.md):
    N: eBay 搜索日期    — YYYY-MM-DD
    O: eBay 同类低价    — numeric (rank-1 or rank-2, whichever is cheaper)
    P: 低价链接         — cleaned eBay itm URL
    Q: eBay 同类高价    — numeric (the pricier of the two)
    R: 高价链接         — cleaned eBay itm URL

Row selection: G exactly 'Done' AND N empty AND K (eBay 标题) non-empty.
"""

import argparse
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

# Column indices for the eBay result block.
COL_EBAY_SEARCH_DATE = 14
COL_LOW_PRICE        = 15
COL_LOW_URL          = 16
COL_HIGH_PRICE       = 17
COL_HIGH_URL         = 18

logger = logging.getLogger("ebay_price_check")
DEBUG_LOG_DIR = Path(__file__).resolve().parent / "debug_logs"


def _read_ebay_pending_rows(ws) -> list:
    """Return list of dicts for rows that need an eBay search:
    G == 'Done' AND N empty AND K (eBay 标题) non-empty. Non-template
    sheets (missing '链接' in B1) return [].
    Dict shape: {row, seq, ebay_title, shein_title, image_col_h_present}."""
    # Import here to avoid circular deps at module load time.
    from run_excel import _sheet_matches_template, COL_SEQ, COL_STATUS, COL_SHEIN_TITLE, COL_EBAY_TITLE
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        status = ws.cell(r, COL_STATUS).value
        search_date = ws.cell(r, COL_EBAY_SEARCH_DATE).value
        ebay_title = ws.cell(r, COL_EBAY_TITLE).value
        shein_title = ws.cell(r, COL_SHEIN_TITLE).value
        if str(status or "") != "Done":
            continue
        if search_date not in (None, ""):
            continue
        if not (ebay_title and str(ebay_title).strip()):
            logger.info("  row %d: skip (eBay 标题 empty)", r)
            continue
        try:
            seq_int = int(seq) if seq is not None else None
        except (TypeError, ValueError):
            seq_int = None
        pending.append({
            "row": r,
            "seq": seq_int,
            "ebay_title": str(ebay_title).strip(),
            "shein_title": str(shein_title or "").strip(),
        })
    return pending


def _write_ebay_result(
    ws, row: int, search_date: str,
    low_price=None, low_url: str = None,
    high_price=None, high_url: str = None,
    no_match: bool = False,
) -> None:
    """Write columns N-R for one row. Always writes N (search_date).
    On no_match=True: O = 'no match', P/Q/R empty. Otherwise: fill the
    slots provided; unspecified stays None."""
    ws.cell(row, COL_EBAY_SEARCH_DATE).value = search_date
    if no_match:
        ws.cell(row, COL_LOW_PRICE).value = "no match"
        return
    if low_price is not None:
        ws.cell(row, COL_LOW_PRICE).value = low_price
    if low_url:
        ws.cell(row, COL_LOW_URL).value = low_url
    if high_price is not None:
        ws.cell(row, COL_HIGH_PRICE).value = high_price
    if high_url:
        ws.cell(row, COL_HIGH_URL).value = high_url
```

- [ ] **Step 4: Run — expect PASS**

```powershell
python test_ebay_io.py
```

Expected: `ALL PASS`.

- [ ] **Step 5: Commit**

```powershell
git add test_ebay_io.py ebay_price_check.py
git commit -m "feat(ebay-check): row reader + writer for template N-R columns"
```

---

## Task 7: Driver — `process_excel` + `main` + 3-tab parallelism

**Files:**
- Modify: `ebay_price_check.py`

- [ ] **Step 1: Extend `ebay_price_check.py` with the driver and CLI**

Append to `C:\Users\ak\Desktop\Claude\shein-extract-au\ebay_price_check.py`:

```python
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from shein_scraper import (
    CDP_PORT,
    MAX_PARALLEL_TABS,
    INTER_URL_DELAY_SEC,
    RateLimitError,
    _ensure_chrome,
    _inter_url_pause,
    _rate_limit_lock,
    _record_result_for_rate_limit,
    _reset_rate_limit_state,
    _screenshots_dir,
    _take_screenshot,
    _ws_url_for_id,
)
from notify import alert_captcha, alert_generic
from ebay_scraper import search_ebay_au, _ensure_ebay_session
from config import SUBMITTED_DIR, INPUT_FILENAME


def setup_logging():
    DEBUG_LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = DEBUG_LOG_DIR / f"ebay_{ts}.log"
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    root = logging.getLogger("ebay_price_check")
    root.handlers.clear()
    root.setLevel(logging.DEBUG)

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    sh.setFormatter(fmt)
    root.addHandler(sh)

    logger.info("Log: %s", log_path)
    return log_path


def safe_save(wb, xlsx_path: Path) -> None:
    """Save workbook. If locked by another user, save as copy with '2' suffix."""
    try:
        wb.save(xlsx_path)
    except PermissionError:
        alt = xlsx_path.with_stem(xlsx_path.stem + "2")
        logger.warning("Cannot save to %s (locked), saving to %s", xlsx_path.name, alt.name)
        wb.save(alt)
        logger.info("Saved to alternate: %s", alt.name)


def _check_one_row(pending_row: dict, port: int) -> dict:
    """Worker: search eBay for one pending row. Returns a result dict with:
      status: 'ok' | 'no_match' | 'captcha' | 'error'
      hits: list[EbayHit] (may be empty)
      row: int (original row for write-back)
      seq: int | None
      error: str | None (populated on captcha/error)
    """
    result = {"row": pending_row["row"], "seq": pending_row.get("seq"),
              "status": "error", "hits": [], "error": None}
    query = pending_row["ebay_title"]
    print(f"[eBay] row {pending_row['row']} seq {pending_row['seq']}: "
          f"searching '{query[:60]}...'")
    try:
        hits = search_ebay_au(query, port)
        if not hits:
            result["status"] = "no_match"
            print("  0 results")
        else:
            result["hits"] = hits
            result["status"] = "ok"
            print(f"  {len(hits)} hit(s): "
                  + ", ".join(f"${h.delivered_price:.2f}" for h in hits))
    except RuntimeError as e:
        msg = str(e)
        if "blocked" in msg.lower() or "captcha" in msg.lower():
            result["status"] = "captcha"
            result["error"] = msg
            print(f"  CAPTCHA: {msg}")
        else:
            result["error"] = msg
            print(f"  ERROR: {msg}")
    except Exception as e:
        result["error"] = str(e)
        print(f"  ERROR: {e}")
    return result


def _apply_result_to_row(ws, result: dict, today: str, ws_lock: threading.Lock) -> None:
    """Serialize the write-back (Excel isn't thread-safe)."""
    with ws_lock:
        row = result["row"]
        status = result["status"]
        if status == "ok":
            hits = result["hits"]
            # Sort by delivered_price: cheaper → O/P, pricier → Q/R.
            # Tie-break stable (Best Match rank 1 wins the O slot).
            sorted_hits = sorted(hits, key=lambda h: h.delivered_price)
            low = sorted_hits[0]
            high = sorted_hits[1] if len(sorted_hits) >= 2 else None
            _write_ebay_result(
                ws, row=row, search_date=today,
                low_price=low.delivered_price, low_url=low.url,
                high_price=(high.delivered_price if high else None),
                high_url=(high.url if high else None),
            )
        elif status == "no_match":
            _write_ebay_result(ws, row=row, search_date=today, no_match=True)
        else:
            # captcha / error → do NOT write N; row retries next run.
            logger.info("    row %d skipped (%s)", row, status)


def process_excel(xlsx_path: Path) -> None:
    """Process every worksheet in the file. See module docstring."""
    logger.info("Opening: %s", xlsx_path.name)
    wb = load_workbook(xlsx_path)
    ws_lock = threading.Lock()

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        pending = _read_ebay_pending_rows(ws)
        if not pending:
            logger.info("  Sheet '%s': no rows to search (or not template schema)",
                        ws_name.strip())
            continue

        logger.info("  Sheet '%s': %d pending row(s): seq %s",
                    ws_name.strip(), len(pending),
                    [p["seq"] for p in pending])

        _reset_rate_limit_state()
        _ensure_ebay_session(CDP_PORT)

        today = datetime.now().strftime("%Y-%m-%d")
        rate_limited = False

        total = len(pending)
        for batch_start in range(0, total, MAX_PARALLEL_TABS):
            batch_end = min(batch_start + MAX_PARALLEL_TABS, total)
            with _rate_limit_lock:
                if rate_limited:
                    break

            with ThreadPoolExecutor(max_workers=MAX_PARALLEL_TABS) as ex:
                futures = [
                    ex.submit(_check_one_row, pending[i], CDP_PORT)
                    for i in range(batch_start, batch_end)
                ]
                tripped_early = False
                for fut in as_completed(futures):
                    try:
                        result = fut.result()
                    except Exception as e:
                        result = {"row": -1, "seq": None, "status": "error",
                                  "hits": [], "error": str(e)}
                    _apply_result_to_row(ws, result, today, ws_lock)
                    if _record_result_for_rate_limit(
                        "OK" if result["status"] in ("ok", "no_match") else "FAIL"
                    ):
                        tripped_early = True
                        break
                # Drain completed-but-uncollected futures.
                if tripped_early:
                    for fut in futures:
                        if fut.done():
                            try:
                                result = fut.result()
                                _apply_result_to_row(ws, result, today, ws_lock)
                            except Exception:
                                pass
                    rate_limited = True

            if batch_end < total and not rate_limited:
                _inter_url_pause(batch_end, total)

        safe_save(wb, xlsx_path)
        logger.info("  Saved progress to %s", xlsx_path.name)

    wb.close()
    logger.info("Done: %s", xlsx_path.name)


def main():
    parser = argparse.ArgumentParser(
        description="Post-Shein-scrape eBay price check helper (澳洲站)")
    parser.add_argument("file", nargs="?", default=None,
                        help="Path to .xlsx (default: SHEIN_INPUT_FILENAME under SUBMITTED_DIR)")
    args = parser.parse_args()

    setup_logging()

    if args.file:
        xlsx_path = Path(args.file)
    elif INPUT_FILENAME:
        xlsx_path = SUBMITTED_DIR / INPUT_FILENAME
    else:
        logger.error("No file given and SHEIN_INPUT_FILENAME not set in .env")
        sys.exit(1)

    if not xlsx_path.exists():
        logger.error("File not found: %s", xlsx_path)
        sys.exit(1)

    _ensure_chrome()  # launches or reuses Chrome on CDP_PORT

    try:
        process_excel(xlsx_path)
    except RateLimitError:
        logger.warning("[限流] Rate limited — some rows left unsearched")
    except Exception as e:
        logger.exception("Fatal error: %s", e)
        try:
            tb = traceback.format_exc()
            (DEBUG_LOG_DIR / "last_traceback.txt").write_text(tb, encoding="utf-8")
        except OSError:
            pass
    logger.info("All done.")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Quick import + signature check**

```powershell
python -c "import ebay_price_check; print('ok'); print(dir(ebay_price_check))" | Select-String -Pattern "process_excel|main|_check_one_row" | Select-Object -First 5
```

Expected: prints `ok` then function names appear in the dir output.

- [ ] **Step 3: Re-run all tests**

```powershell
python test_variant_merge.py
python test_variant_filter.py
python test_price_stock_format.py
python test_template_io.py
python test_ebay_url_clean.py
python test_ebay_parse.py
python test_ebay_extract.py
python test_ebay_io.py
```

Expected: eight `ALL PASS`.

- [ ] **Step 4: Commit**

```powershell
git add ebay_price_check.py
git commit -m "feat(ebay-check): process_excel driver with 3-tab parallelism + main CLI"
```

---

## Task 8: End-to-end manual smoke against real ZR file

**Files:** none new — verifies the end-to-end pipeline against real eBay AU.

- [ ] **Step 1: Clear N-R for the 8 seeded rows in the ZR file**

Run in bash (double-quoted `python -c` outer, single-quoted Python literals inner):

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook
p = r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接汇总 - ZR.xlsx'
wb = load_workbook(p)
ws = wb['ZR1']
for r in range(2, 10):
    for c in range(14, 19):
        ws.cell(r, c).value = None
wb.save(p)
print('cleared N-R for rows 2-9')
"
```

- [ ] **Step 2: Run the check**

```powershell
python -X utf8 ebay_price_check.py
```

Expect the log to show:
- Session warmup: `[导航] 预热：先访问 ebay.com.au 首页建立 session...`
- 8 pending rows detected: `Sheet 'ZR1': 8 pending row(s): seq [1, 2, 3, 4, 5, 6, 7, 8]`
- 3-tab parallel batches with `[节奏] 间隔 3s` between them
- Per-row lines: `[eBay] row N seq S: searching 'xxx'` followed by `2 hit(s): $A, $B`
- Wall-clock 30–90 seconds total

- [ ] **Step 3: Verify results in the file**

```bash
python -X utf8 -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook
p = r'D:\共享云端硬盘\02 希音\01 店铺资料\99 ZR\澳洲希音链接汇总 - ZR.xlsx'
wb = load_workbook(p, data_only=True)
ws = wb['ZR1']
print(f'{\"seq\":<4} {\"N搜索日期\":<12} {\"O低价\":<10} {\"Q高价\":<10} 链接示例')
for r in range(2, 10):
    seq = ws.cell(r, 1).value
    n = ws.cell(r, 14).value
    o = ws.cell(r, 15).value
    q = ws.cell(r, 17).value
    p_url = ws.cell(r, 16).value
    print(f'{seq!s:<4} {n!s:<12} {o!s:<10} {q!s:<10} {p_url}')
"
```

Expect:
- Every row has a date in N.
- O column is numeric (float) or the string `no match`.
- Q column is numeric or blank (only 1 hit).
- P/R are `https://www.ebay.com.au/itm/...` URLs with no query string.

- [ ] **Step 4: Manual eyeball check — click 2 or 3 P/R links**

Open a couple of the eBay URLs in a real browser. Confirm each is:
- A currently-listed product (not a dead link).
- Roughly the same product category as the row's 希音标题 / eBay标题.

If most rows produce plausible matches, the pipeline works. Widespread "no match" or clearly-wrong matches indicates either eBay's DOM changed (fix the extraction JS selectors) or the query needs refinement (widen search filters).

- [ ] **Step 5: If smoke passes, no commit needed — this task is verification only.**

If issues found, file them as follow-ups before merging the branch.

---

## Rollout

- Merge `feat/ebay-price-check` to `main` after Task 8 passes.
- No installer bump strictly required unless you plan to ship this to employees as part of the exe (`ebay_price_check.py` is a separate entry point — not currently reachable from `app_main.py`'s menu). If you want it in the compiled exe, that's a follow-on task (add a menu entry in `app_main.py`, add `ebay_price_check` and `ebay_scraper` to `pyinstaller.spec` hidden imports, bump version, `build.bat`).
- Otherwise the script runs from source on the owner's machine only, which matches the manual 2审 use case.
