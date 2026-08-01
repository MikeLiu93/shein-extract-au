"""Tests for _format_price_range (H column '希音价格') and _format_stock_summary
(L column '库存') — Chinese-header template output columns."""
from shein_scraper import _format_price_range, _format_stock_summary


# ── _format_price_range ─────────────────────────────────────────────────────

def test_empty_returns_none():
    assert _format_price_range([]) is None


def test_single_price_returns_numeric():
    skus = [{"sale_price": 9.99}]
    assert _format_price_range(skus) == 9.99


def test_all_same_price_returns_numeric():
    skus = [{"sale_price": 9.99}, {"sale_price": 9.99}, {"sale_price": 9.99}]
    assert _format_price_range(skus) == 9.99


def test_differing_prices_return_range_string_no_currency():
    skus = [{"sale_price": 9.99}, {"sale_price": 14.99}, {"sale_price": 12.50}]
    # En-dash U+2013, no '$' — user wants numeric where possible, plain range otherwise.
    assert _format_price_range(skus) == "9.99–14.99"


def test_ignores_none_prices():
    skus = [{"sale_price": None}, {"sale_price": 9.99}]
    assert _format_price_range(skus) == 9.99


def test_only_nones_returns_none():
    skus = [{"sale_price": None}, {"sale_price": None}]
    assert _format_price_range(skus) is None


# ── _format_stock_summary ───────────────────────────────────────────────────

def test_stock_empty_returns_empty():
    assert _format_stock_summary([]) == ""


def test_stock_single_variant_no_prefix():
    skus = [{"attrs": {"Color": "Black"}, "stock": 20}]
    assert _format_stock_summary(skus) == "Black: 20"


def test_stock_multi_variant_slash_separated():
    skus = [
        {"attrs": {"Color": "Black", "Size": "M"}, "stock": 20},
        {"attrs": {"Color": "Black", "Size": "L"}, "stock": 0},
        {"attrs": {"Color": "Red",   "Size": "M"}, "stock": 3},
    ]
    out = _format_stock_summary(skus)
    assert out == "Black-M: 20 / Black-L: 缺货 / Red-M: 少货 3", out


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
    test_empty_returns_none()
    test_single_price_returns_numeric()
    test_all_same_price_returns_numeric()
    test_differing_prices_return_range_string_no_currency()
    test_ignores_none_prices()
    test_only_nones_returns_none()
    test_stock_empty_returns_empty()
    test_stock_single_variant_no_prefix()
    test_stock_multi_variant_slash_separated()
    test_stock_low_stock_threshold_15()
    test_stock_missing_attrs_uses_sku_code()
    print("ALL PASS")
