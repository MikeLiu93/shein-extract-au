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


# Matches "12.99", "1,234.99", "50", "1234567.99" — a digit run possibly
# containing commas, optionally with a decimal. Comma placement isn't
# validated (we strip commas before float()); focus is defensive
# extraction, not format conformance.
_PRICE_NUM_RE = re.compile(r"(\d[\d,]*(?:\.\d{1,2})?)")


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


# ── JavaScript: extract top search results from ebay.com.au SRP ─────────────
# eBay rolled out `.s-card` in late 2026 (new "su-*" design system) — the
# older `.s-item` selectors are gone. Placeholder cards with fake test URLs
# (e.g. `ebay.com/itm/123456`) sometimes appear as the first 1-2 items; we
# filter by requiring a real `www.ebay.com.au/itm/<numeric-id>` href.
_JS_EBAY_SEARCH_EXTRACT = r"""
(function() {
    const cards = document.querySelectorAll('li.s-card');
    const out = [];
    for (const c of cards) {
        // The listing link. Real eBay AU items have www.ebay.com.au/itm/<id>.
        const linkEl = c.querySelector('a.s-card__link, a[href*="/itm/"]');
        const url = linkEl?.href || '';
        // Skip placeholder cards ('ebay.com/itm/123456' fake URLs).
        if (!/\bebay\.com\.au\/itm\/\d/.test(url)) continue;

        const t = (c.querySelector('.s-card__title, [class*="s-card__title"]')?.innerText || '').trim();
        if (!t || /^shop on ebay$/i.test(t)) continue;

        // Price lives in .s-card__price (usually inside a `su-styled-text` span).
        const priceEl = c.querySelector('.s-card__price, [class*="s-card__price"]');
        const price = (priceEl?.innerText || '').trim();

        // Postage: eBay AU uses .s-card__subtitle-row / .s-card__subtitle
        // for shipping labels. Fall back to any nearby element mentioning
        // postage/shipping/delivery.
        const postEl =
            c.querySelector('.s-card__subtitle-row, .s-card__subtitle') ||
            [...c.querySelectorAll('span,div')].find(e =>
                /postage|shipping|delivery/i.test(e.innerText || ''));
        const postage = (postEl?.innerText || '').trim();

        const imgEl = c.querySelector('img[src*="ebayimg"], img.s-card__image');
        const image_url = imgEl?.src || '';

        if (t && price && url) {
            out.push({title: t, price: price, postage: postage, url: url, image_url: image_url});
        }
        if (out.length >= 5) break;
    }
    // Signal whether the SRP shell has rendered (page loaded, ready to inspect).
    // If the shell is absent, caller treats this as a timeout, not a genuine 0-result.
    var loaded = !!document.querySelector(
        '#srp-river-results, .srp-results, .srp-controls, .srp-river-main, [class*="srp-list"]'
    );
    return {loaded: loaded, hits: out};
})()
"""


def _shorten_query(k) -> str:
    """Turn an 80-char AI-optimized eBay listing title into a 3-5 word search
    query. AI titles are keyword-packed for SEO discovery on eBay LISTINGS,
    but as SEARCH queries they return 0 results because eBay treats the
    long string as a strict phrase match. Strips leading quantity tokens
    ('12pcs', '100Pcs', '1/2pcs', '1 Set') to focus on product nouns, then
    caps at 5 words / 50 chars."""
    if not k or not isinstance(k, str):
        return ""
    words = k.strip().split()
    # Strip leading quantity tokens (up to 2 at the front).
    # Matches: 12pcs, 100Pcs, 1/2pcs, 1, 2Pcs, 1Set, 1-2pcs, etc.
    qty_pat = re.compile(r'^\d[\dxX/\-]*(pc|pcs|set|sets|pack|packs)?$', re.IGNORECASE)
    setword_pat = re.compile(r'^(pc|pcs|set|sets|pack|packs|piece|pieces)$', re.IGNORECASE)
    dropped = 0
    while words and dropped < 2:
        if qty_pat.match(words[0]) or setword_pat.match(words[0]):
            words.pop(0)
            dropped += 1
        else:
            break
    picked = words[:5]
    q = " ".join(picked)
    if len(q) > 50:
        cut = q[:50].rsplit(" ", 1)[0]
        q = cut
    return q


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
    results. Raises RuntimeError on captcha (caller handles retry).

    `query` is expected to be the K-column AI-optimized eBay listing title.
    It gets internally shortened via _shorten_query — eBay treats 80-char
    strings as a phrase match and returns 0 results."""
    from urllib.parse import quote
    from shein_scraper import _JS_DETECT_BLOCK  # reuse existing block detector

    short = _shorten_query(query)
    if not short:
        return []
    print(f"    [query] '{query[:50]}...' -> '{short}'")
    q = quote(short, safe="")
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
        page_loaded = False
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
                result = _run_js(ws_url, _JS_EBAY_SEARCH_EXTRACT) or {}
                if isinstance(result, dict):
                    page_loaded = page_loaded or bool(result.get("loaded"))
                    raw_hits = result.get("hits") or []
                    if raw_hits:
                        break
                    if page_loaded:
                        # Page rendered but no hits — genuine 0-result. Stop polling.
                        break
            except Exception:
                pass
            time.sleep(1.0)
        if not page_loaded and not raw_hits:
            # SRP shell never rendered — treat as timeout (retryable, not "no match").
            raise RuntimeError("eBay search timeout: SRP shell never rendered")
        return _extract_from_page(raw_hits)
    finally:
        _close_tab(port, tab_id)
