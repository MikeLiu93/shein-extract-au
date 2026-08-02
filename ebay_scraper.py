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
