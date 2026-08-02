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


def test_price_unformatted_large_number():
    # Regression: no commas, 4+ digits — must not silently truncate.
    assert _parse_price("$1234567.99") == 1234567.99
    assert _parse_price("AU $9999") == 9999.0


if __name__ == "__main__":
    test_price_au_prefix()
    test_price_bare_dollar()
    test_price_thousand_separator()
    test_price_integer_no_decimals()
    test_price_range_uses_low_end()
    test_price_whitespace_tolerant()
    test_price_empty_returns_none()
    test_price_gibberish_returns_none()
    test_price_unformatted_large_number()
    test_postage_free()
    test_postage_free_case_insensitive()
    test_postage_plus_prefix()
    test_postage_no_plus()
    test_postage_empty_returns_none()
    test_postage_unparseable_returns_none()
    print("ALL PASS")
