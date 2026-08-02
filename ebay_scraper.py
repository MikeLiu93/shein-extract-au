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
