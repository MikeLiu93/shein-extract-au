"""Tests for the tiered eBay pricing formula (2026-09).

Rule:
    price <  LOW_PRICE_THRESHOLD (20)  →  price + LOW_PRICE_FLAT_MARKUP (10) + shipping
    price >= LOW_PRICE_THRESHOLD       →  price × EBAY_MARKUP               + shipping

Default markup at test time is 1.2 (config default; wizard-configurable).
"""
from shein_scraper import (
    _ebay_listing_price,
    EBAY_MARKUP, LOW_PRICE_THRESHOLD, LOW_PRICE_FLAT_MARKUP,
)


# ── Low-price branch (flat +$10) ─────────────────────────────────────────────

def test_cheap_item_no_shipping_gets_flat_markup():
    # p=5 → 5+10+0 = 15
    assert _ebay_listing_price(5, 0) == 5 + LOW_PRICE_FLAT_MARKUP


def test_cheap_item_with_shipping_adds_both():
    # p=5, s=7.95 → 5+10+7.95 = 22.95
    assert _ebay_listing_price(5, 7.95) == round(5 + LOW_PRICE_FLAT_MARKUP + 7.95, 2)


def test_just_below_threshold_uses_flat():
    # p=19.99 → 19.99+10+ship  (NOT 19.99×1.2)
    got = _ebay_listing_price(19.99, 0)
    assert got == round(19.99 + LOW_PRICE_FLAT_MARKUP, 2), got


# ── Boundary at threshold ────────────────────────────────────────────────────

def test_at_threshold_uses_ratio_branch():
    # p=20 → 20×MARKUP+ship (NOT 20+10+ship)
    got = _ebay_listing_price(LOW_PRICE_THRESHOLD, 0)
    assert got == round(LOW_PRICE_THRESHOLD * EBAY_MARKUP, 2), got


def test_boundary_discontinuity_ratio_lower_than_flat():
    """At markup 1.2 the ratio price at $20 (=$24) is deliberately less than
    the flat price at $19.99 (=$29.99). Business chose this — don't smooth it
    accidentally in future refactors."""
    if EBAY_MARKUP >= 1.5:
        return   # test only meaningful for low-markup configs
    flat_at_1999 = _ebay_listing_price(19.99, 0)
    ratio_at_20 = _ebay_listing_price(20.00, 0)
    assert flat_at_1999 > ratio_at_20, (flat_at_1999, ratio_at_20)


# ── High-price branch (× markup) ─────────────────────────────────────────────

def test_expensive_item_uses_ratio():
    # p=100 → 100×MARKUP+ship
    got = _ebay_listing_price(100, 0)
    assert got == round(100 * EBAY_MARKUP, 2), got


def test_expensive_item_with_shipping():
    got = _ebay_listing_price(50, 3.5)
    assert got == round(50 * EBAY_MARKUP + 3.5, 2), got


# ── Edge / defensive ─────────────────────────────────────────────────────────

def test_none_price_treated_as_zero_hits_low_branch():
    # p=None → 0, 0+10+0 = 10
    assert _ebay_listing_price(None, 0) == LOW_PRICE_FLAT_MARKUP


def test_none_shipping_treated_as_zero():
    assert _ebay_listing_price(30, None) == round(30 * EBAY_MARKUP, 2)


def test_string_inputs_coerce():
    assert _ebay_listing_price("15", "2") == round(15 + LOW_PRICE_FLAT_MARKUP + 2, 2)
    assert _ebay_listing_price("25", "0") == round(25 * EBAY_MARKUP, 2)


def test_result_is_rounded_to_2dp():
    got = _ebay_listing_price(10.001, 0)
    # Any weird fp math should still round cleanly
    assert isinstance(got, float)
    assert round(got, 2) == got


if __name__ == "__main__":
    test_cheap_item_no_shipping_gets_flat_markup()
    test_cheap_item_with_shipping_adds_both()
    test_just_below_threshold_uses_flat()
    test_at_threshold_uses_ratio_branch()
    test_boundary_discontinuity_ratio_lower_than_flat()
    test_expensive_item_uses_ratio()
    test_expensive_item_with_shipping()
    test_none_price_treated_as_zero_hits_low_branch()
    test_none_shipping_treated_as_zero()
    test_string_inputs_coerce()
    test_result_is_rounded_to_2dp()
    print("ALL PASS")
