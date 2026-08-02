# eBay Price Check — Design

**Date:** 2026-08-02
**Scope:** `shein-extract-au` (澳洲站)
**Status:** Approved for implementation planning.
**Follow-on to:** `2026-08-01-au-template-refactor-design.md` (template schema).

## Purpose

After `run_excel.py` finishes scraping Shein and filling F–M of a row, an operator needs to sanity-check the K-column eBay price (`(原价 or 希音价) × 2 + 运费`) against what actually exists on `ebay.com.au`. Manual check for 8–100 rows/day is tedious and error-prone.

This new script — `ebay_price_check.py` — is a *post-scrape helper*. For each ready row it searches `ebay.com.au` for the same product using the AI-generated K-column title, extracts the top 2 Best Match listings' delivered prices + URLs, and writes them to columns N–R. The operator eyeballs O vs L (my price) and Q vs L to decide whether to publish, raise, or lower.

The script **does not decide pricing**. It surfaces market data for human judgement — same philosophy as I-column 希音价格 (check, not decide).

## Success Criteria

- Row with `G=Done` and `N=empty` gets `N` (search date), `O` (rank-1 delivered price), `P` (rank-1 clean URL), `Q` (rank-2 delivered price), `R` (rank-2 clean URL) filled after one script invocation.
- Filled rows are skipped on subsequent runs (idempotent). Re-search by clearing N.
- Operator can point at O + Q and answer "is my L reasonable?" in under 5 seconds per row.
- Zero manual eBay searches for the common case.

## §1 — Trigger & Row Selection

- CLI: `python ebay_price_check.py [<xlsx_path>]`.
- No path → uses `SHEIN_INPUT_FILENAME` from `.env` (currently `澳洲希音链接汇总 - ZR.xlsx`) resolved against `SHEIN_SUBMITTED_DIR`.
- For each worksheet whose header row matches the template (`B1 == '链接'`), pick rows where:
  - `G (状态)` **exactly equals** `Done` (case-sensitive; `Failed` / `Delisted` / any other value → skip), AND
  - `N (eBay 搜索日期)` is empty
- Non-template sheets skipped silently (same rule as `run_excel.py`).
- Re-run policy: clear N (and O–R if you want a clean re-write). Idempotent — no double-search.
- Session warmup (`https://www.ebay.com.au/` once per process invocation, gated by an `_EBAY_SESSION_WARMED` module flag) runs before the ThreadPoolExecutor pool starts.

## §2 — Search URL & Filters

Query = **K column verbatim**, URL-encoded. K is the AI-generated eBay-optimized title (≤80 chars) — already tuned for eBay search relevance.

```
https://www.ebay.com.au/sch/i.html
  ?_nkw=<url-encoded K>
  &LH_BIN=1                # Buy It Now (exclude auctions)
  &LH_ItemCondition=1000   # New condition only
  &LH_PrefLoc=1            # Located in Australia
  &_sop=12                 # Best Match sort (explicit; guards against eBay changing default)
```

Rationale for filters: comparable pricing benchmark requires competing on the same terms as an AU-based new-item listing. Used items and international sellers use different pricing physics.

## §3 — Result Extraction

- Wait for `.srp-results` grid (or equivalent) to appear.
- For each of the first 2 result cards, extract:
  - `title` — visible listing title
  - `sticker_price` — the highlighted price (AUD)
  - `postage` — the "+ $X postage" line (or "Free postage" → 0)
  - `url` — anchor href on the title
  - `image_url` — first-thumbnail src (recorded for future image-verification phase; not written)
- **Delivered price** = `sticker_price + postage`. If postage parse fails, fall back to `sticker_price` alone and log a warning (row still filled).
- **Ranking**: eBay's Best Match order (rank 1, rank 2).
- **Slot assignment**: sort the two ranked candidates by delivered price → cheaper goes to O/P (`低价`), pricier to Q/R (`高价`). Preserves the column names' intent while honouring "top 2 similar" as the pool. Tie-break on equal delivered price: Best Match rank 1 wins the O slot.
- **URL cleanup**: strip the entire query string from an eBay listing URL. Result form: `https://www.ebay.com.au/itm/<numeric-id>` (or `https://www.ebay.com.au/itm/<slug>/<numeric-id>` if the slug variant appears). Retain path + host + scheme; drop everything after `?`. Non-eBay URLs (shouldn't happen but defensive) pass through unchanged.

## §4 — Write-Back Semantics

| Col | Type | On success (2 results) | On 1 result | On 0 results | On error (network/timeout) |
|-----|------|------------------------|-------------|--------------|----------------------------|
| N (搜索日期) | str `YYYY-MM-DD` | today | today | today | **not written** (retry next run) |
| O (低价) | float | cheaper delivered price | that one price | `"no match"` (string) | not written |
| P (低价链接) | str | cleaned URL | that one URL | empty | not written |
| Q (高价) | float | pricier delivered price | empty | empty | not written |
| R (高价链接) | str | cleaned URL | empty | empty | not written |

Zero results still marks N — a legitimate "we searched, nothing came up" state that shouldn't retry endlessly. Errors leave N empty so the row auto-retries. Captcha alerts via `notify.alert_captcha` and leaves N empty.

Edge — K column empty (Shein didn't generate an eBay title, or the operator cleared it): log info, skip row, N stays empty.

## §5 — Concurrency & Reuse

Mirrors the Shein scraper's post-Task-10 architecture:

- `ThreadPoolExecutor(max_workers=MAX_PARALLEL_TABS)` — 3 concurrent tabs.
- `INTER_URL_DELAY_SEC` — 3 seconds between batches.
- Shared thread-safe rate-limit counter (`_rate_limit_lock`, `_record_result_for_rate_limit`) — 3 consecutive failures → stop batch.
- Session warmup: navigate to `https://www.ebay.com.au/` once before the pool starts (mirrors `_ensure_shein_session`). Set an `_EBAY_SESSION_WARMED` module flag.
- Each worker opens its own tab via `_new_tab`, closes it in `finally`.
- Excel writes serialized behind a `threading.Lock` (same pattern as run_excel).

Reused verbatim from `shein_scraper.py`: `_ensure_chrome`, `_new_tab`, `_close_tab`, `_ws_url_for_id`, `_run_js`, `_cdp_once`, `_check_and_handle_block`, `_take_screenshot`, `_screenshots_dir`, `_inter_url_pause`. Constants: `CDP_PORT`, `MAX_PARALLEL_TABS`, `INTER_URL_DELAY_SEC`, `RATE_LIMIT_CONSECUTIVE`.

**eBay-specific captcha detection**: `_check_and_handle_block` targets Shein's sign-in modal + geetest selectors. eBay uses different DOM (`captcha` iframe from Arkose Labs, sign-in redirect). Extend `_check_and_handle_block` to include eBay selectors, OR add a small `_check_and_handle_ebay_block` wrapper. Decision: extend the shared function with additional selectors — fewer moving parts, same detection philosophy.

## §6 — File Structure

**New files:**

| File | Responsibility |
|------|----------------|
| `ebay_price_check.py` | CLI entry, row loop, `process_excel()`-equivalent driver, per-row write-back |
| `ebay_scraper.py` | `search_ebay_au(query, tab_id, ws_url) -> list[EbayHit]`, JS extraction payload, URL cleanup, price/postage parsing |
| `test_ebay_url_clean.py` | Pure-function tests for the URL cleaner |
| `test_ebay_parse.py` | Pure-function tests for postage/price parsing (given raw strings like "AU $12.99", "+ AU $6.95 postage", "Free postage") |

**Why split from `shein_scraper.py`?** `shein_scraper.py` is already ~3100 lines. eBay search is a different domain (listing grid vs product detail page) with its own JS payload. Independent module → easier to test, understand, and modify without risking Shein regressions.

**Modified files:**

- `shein_scraper.py` — extend `_check_and_handle_block` with eBay captcha/sign-in selectors. Or add a small parallel function. **Decision:** extend existing (see §5).

**No changes to `run_excel.py`.** eBay check is an independent post-processing pass.

## §7 — Data Contract

```python
# ebay_scraper.py
from dataclasses import dataclass

@dataclass
class EbayHit:
    title: str
    sticker_price: float        # AUD, from the highlighted price
    postage: float              # AUD, 0.0 for "Free postage"; None if unparseable
    delivered_price: float      # sticker + postage; falls back to sticker if postage is None
    url: str                    # cleaned (no tracking params)
    image_url: str              # first thumbnail; not written to xlsx (phase 2)

def search_ebay_au(query: str, port: int) -> list[EbayHit]:
    """Open a fresh tab, run the search, return up to 2 hits (Best Match order).
    Raises RuntimeError on captcha; returns [] on zero-result page."""
```

## §8 — Failure Modes

| Signal | Action |
|--------|--------|
| Search page 404 or eBay outage | Row skipped, N not written, log error, retry next run |
| Captcha detected (Arkose iframe, "Please verify" page) | Screenshot, `alert_captcha`, row skipped, N not written |
| Sign-in prompt overlay | Attempt to dismiss (Escape key + close-button click) same as Shein |
| Zero results on the page | N written with today's date, O=`"no match"`, others empty. No retry. |
| Price/postage regex fails on 1 of 2 hits | Use sticker only for that hit, log warning, row still written |
| Rate-limit trip (3 consecutive failures) | Stop current sheet, log warning, retry next invocation |

## §9 — Testing

**Unit** (plain-`assert`, no pytest — match `test_variant_merge.py` convention):

- `test_ebay_url_clean.py`
  - Strip `?campid=...`
  - Strip `?hash=...`
  - Strip everything after `/itm/<numeric id>` (except the ID itself)
  - Preserve URL scheme + host
  - Passthrough for already-clean URLs
- `test_ebay_parse.py`
  - `"AU $12.99"` → 12.99
  - `"AU $1,234.99"` → 1234.99
  - `"$12.99"` → 12.99
  - `"Free postage"` → 0.0
  - `"+ AU $6.95 postage"` → 6.95
  - `""` → None (unparseable)
  - Range price like `"AU $12.99 to AU $18.99"` — decision: take the low end

**Manual smoke** (before merge):

- Run against the ZR template after Shein has filled 8 rows.
- Expect 8 rows to gain N/O/P/Q/R data; scan for one Q that's suspiciously higher than O (Shein-copy listings often cluster tightly, so real spread implies a bad match — good signal).
- Verify N is a date string; O and Q are numeric (Excel formats as numbers, not strings); P and R are clickable URLs.
- Verify wall-clock: 8 rows × ~10s per row / 3-tab parallelism ≈ 30–60 seconds total.

## §10 — Out of Scope

- Image similarity verification (phase 2 candidate).
- Multi-marketplace search (only `ebay.com.au`).
- Auction listings (BIN filter excludes them).
- Used items.
- Automatic pricing recommendation (`不加，YAGNI`).
- Historical price tracking (would need a separate log table).
- Currency conversion (AU only, always AUD).
- Extending to non-ZR sheets — logic is sheet-agnostic; works as long as headers match.

## §11 — Migration & Risk

- Backwards-compatible with the existing template — only fills unused columns.
- Zero effect on `run_excel.py` behaviour or the batch xlsx.
- **Risk**: eBay's search-result DOM changes periodically. Mitigation: extraction JS uses stable selectors + defensive fallback (returns partial data rather than exceptions); test-fixture updates when the operator reports a "wrong price" incident.
- **Risk**: eBay may block automated searches at scale. Mitigation: 3-tab conservative rate; per-row session warmup; captcha alert path already in place.
