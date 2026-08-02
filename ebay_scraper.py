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
