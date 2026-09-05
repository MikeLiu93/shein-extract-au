"""Tests for the new price-based shipping rule.

Rule: price >= FREE_SHIPPING_THRESHOLD (default 9) → free; else DEFAULT_SHIPPING_FEE (7.95).
"""
from shein_scraper import _calc_shipping, FREE_SHIPPING_THRESHOLD, DEFAULT_SHIPPING_FEE


def test_price_at_threshold_is_free():
    assert _calc_shipping(FREE_SHIPPING_THRESHOLD) == 0.0


def test_price_above_threshold_is_free():
    assert _calc_shipping(FREE_SHIPPING_THRESHOLD + 1) == 0.0


def test_price_below_threshold_pays_default_fee():
    assert _calc_shipping(FREE_SHIPPING_THRESHOLD - 0.01) == DEFAULT_SHIPPING_FEE


def test_price_zero_pays_default_fee():
    assert _calc_shipping(0) == DEFAULT_SHIPPING_FEE


def test_price_none_is_treated_as_zero():
    assert _calc_shipping(None) == DEFAULT_SHIPPING_FEE


def test_string_price_coerces():
    assert _calc_shipping("15") == 0.0
    assert _calc_shipping("3.5") == DEFAULT_SHIPPING_FEE


if __name__ == "__main__":
    test_price_at_threshold_is_free()
    test_price_above_threshold_is_free()
    test_price_below_threshold_pays_default_fee()
    test_price_zero_pays_default_fee()
    test_price_none_is_treated_as_zero()
    test_string_price_coerces()
    print("ALL PASS")
