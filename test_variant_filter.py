"""Tests for _filter_variants_by_declaration — the E-column filter that
restricts sku_prices + variations to only the user's declared variants."""
from shein_scraper import _filter_variants_by_declaration


def _make_sku(sku_code, color, size, price=9.99, stock=10):
    return {
        "sku_code": sku_code,
        "attrs": {"Color": color, "Size": size},
        "sale_price": price,
        "retail_price": price,
        "stock": stock,
    }


SAMPLE_SKUS = [
    _make_sku("A1", "Black", "M", 9.99, 12),
    _make_sku("A2", "Black", "L", 9.99, 0),
    _make_sku("A3", "Red",   "M", 10.99, 3),
    _make_sku("A4", "Red",   "L", 10.99, 20),
    _make_sku("A5", "Blue",  "M", 9.99, 5),
]
SAMPLE_VARS = {"Color": ["Black", "Red", "Blue"], "Size": ["M", "L"]}


def test_empty_declaration_returns_everything():
    kept_skus, kept_vars, unknown = _filter_variants_by_declaration(
        "", SAMPLE_SKUS, SAMPLE_VARS
    )
    assert kept_skus == SAMPLE_SKUS
    assert kept_vars == SAMPLE_VARS
    assert unknown == []


def test_two_groups_cartesian():
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Red / M, L", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A2", "A3", "A4"], codes
    assert vars_ == {"Color": ["Black", "Red"], "Size": ["M", "L"]}
    assert unknown == []


def test_single_group_flat_allowlist():
    # No "/" → allow-list matched against any attribute value.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Red", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A2", "A3", "A4"], codes
    assert vars_["Color"] == ["Black", "Red"]
    assert unknown == []


def test_case_and_whitespace_insensitive():
    kept, vars_, unknown = _filter_variants_by_declaration(
        "  black ,RED /  m ", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A3"], codes
    assert unknown == []


def test_unknown_value_but_partial_match():
    # "Green" doesn't exist; "Black" does. Keep Black rows, report Green.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Green / M", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1"], codes
    assert "Green" in unknown, unknown


def test_zero_match_returns_empty_and_flags():
    # Nothing matches — caller (run_excel) will surface as Failed.
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Purple / XXL", SAMPLE_SKUS, SAMPLE_VARS
    )
    assert kept == []
    assert vars_ == {}
    assert set(unknown) >= {"Purple", "XXL"}


def test_flat_variations_key_not_in_sku_attrs():
    """A variations key with values but no per-SKU attrs entry: the user
    hint enriches the SKU AND filtered_vars carries the matched value
    forward. Cotton is a real product attribute that matched, so it should
    NOT appear in unknown.
    """
    skus = [
        {"sku_code": "X1", "attrs": {"Color": "Black"}, "sale_price": 5.0, "stock": 1},
    ]
    variations = {
        "Color": ["Black", "Red"],
        "Material": ["Cotton", "Wool"],
    }
    # Flat (no "/") declaration: X1 matches (Black in attrs, and Cotton
    # matched via variations enrichment).
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Cotton", skus, variations
    )
    assert [s["sku_code"] for s in kept] == ["X1"]
    # Color rebuilt from X1's attrs.
    assert vars_.get("Color") == ["Black"]
    # Material carried over from the enrichment lookup.
    assert vars_.get("Material") == ["Cotton"]
    # Both Black and Cotton were matched → unknown is empty.
    assert unknown == []


def test_multi_group_seen_leak_reports_unreachable_value():
    """Bug repro: 'Red' declared, but no Red row survives the size filter →
    'Red' MUST appear in unknown so the operator knows it's unreachable."""
    skus = [
        _make_sku("A1", "Black", "M",  9.99, 12),
        _make_sku("A2", "Red",   "XL", 10.99, 3),
    ]
    variations = {"Color": ["Black", "Red"], "Size": ["M", "L", "XL"]}
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Black, Red / M, L", skus, variations
    )
    assert [s["sku_code"] for s in kept] == ["A1"]
    assert "Red" in unknown, f"expected 'Red' unreachable warning; got unknown={unknown}"
    assert "XL" not in unknown, f"XL wasn't declared; got unknown={unknown}"


def test_attribute_prefix_stripped():
    """User writes 'Color:Pink' (attribute-value form) — script must strip
    the 'Color:' prefix and match 'Pink' against SKU attr values."""
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Color:Black, Color:Red", SAMPLE_SKUS, SAMPLE_VARS
    )
    codes = [s["sku_code"] for s in kept]
    assert codes == ["A1", "A2", "A3", "A4"], codes
    assert unknown == []


def test_full_width_colon_prefix_stripped():
    """Chinese full-width colon (：) also acts as attribute separator."""
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Color：Black", SAMPLE_SKUS, SAMPLE_VARS
    )
    assert [s["sku_code"] for s in kept] == ["A1", "A2"]


def test_token_match_substring_pink_matches_1pc_pink():
    """Token-set match: user 'Pink' matches Shein's '1PC Pink' (a common
    Shein naming pattern for single-piece variants)."""
    skus = [
        {"sku_code": "X1", "attrs": {"Color": "1PC Pink"}, "sale_price": 5.0, "stock": 1},
        {"sku_code": "X2", "attrs": {"Color": "1PC Rose Pink"}, "sale_price": 6.0, "stock": 1},
        {"sku_code": "X3", "attrs": {"Color": "1PC Black"}, "sale_price": 5.0, "stock": 1},
    ]
    variations = {"Color": ["1PC Pink", "1PC Rose Pink", "1PC Black"]}
    kept, vars_, unknown = _filter_variants_by_declaration("Pink", skus, variations)
    codes = [s["sku_code"] for s in kept]
    # Both X1 (1PC Pink) and X2 (1PC Rose Pink) contain 'pink' token → both match.
    # X3 (1PC Black) does not.
    assert codes == ["X1", "X2"], codes


def test_token_match_not_pinkeye():
    """Token-set match must NOT be raw substring — 'Pink' should not match
    'Pinkeye' (different tokens, no subset relationship)."""
    skus = [
        {"sku_code": "P1", "attrs": {"Color": "Pinkeye"}, "sale_price": 5.0, "stock": 1},
    ]
    kept, _, _ = _filter_variants_by_declaration("Pink", skus, {"Color": ["Pinkeye"]})
    assert kept == [], "Pink should NOT match Pinkeye (word-boundary safety)"


def test_single_value_attribute_enriched_to_all_skus():
    """AOYI-simple case: product has variations={'Color': ['Pink']} — one
    color, one SKU with only Size attr. User declares 'Color:Pink' — script
    enriches SKU with the implicit single-value Color and matches."""
    skus = [
        {"sku_code": "AOYI-1", "attrs": {"Size": "one-size"},
         "sale_price": 13.25, "stock": 5},
    ]
    variations = {"Color": ["Pink"]}
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Color:Pink", skus, variations
    )
    assert [s["sku_code"] for s in kept] == ["AOYI-1"]
    assert vars_.get("Color") == ["Pink"]
    assert vars_.get("Size") == ["one-size"]
    assert unknown == []


def test_multi_value_top_level_attr_uses_user_hint():
    """AOYI-real case: product's `main_sale_attrs` expands variations to
    {'Color': ['Green', 'Pink', 'White']} — 3 colors, each a separate URL.
    Current URL is Pink (per Shein's main_sale_attrs) but sku_prices only
    has {'Size': 'one-size'}. User declares 'Color:Pink' — script trusts
    the user's hint and injects Color=Pink into the SKU (since 'Pink' IS
    one of the variation values)."""
    skus = [
        {"sku_code": "AOYI-1", "attrs": {"Size": "one-size"},
         "sale_price": 13.25, "stock": 5},
    ]
    variations = {"Color": ["Green", "Pink", "White"]}
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Color:Pink", skus, variations
    )
    assert [s["sku_code"] for s in kept] == ["AOYI-1"], (
        "Should match: user hinted Pink which is in variations")
    assert vars_.get("Color") == ["Pink"]  # only the specific one, not all 3
    assert unknown == []


def test_multi_value_top_level_attr_rejects_when_user_hint_absent():
    """Same AOYI product structure, but user declares a color that's NOT
    in variations. No injection → filter fails → SKU rejected. Correct
    behaviour: don't blindly trust user; require variations to list the
    value they claimed."""
    skus = [
        {"sku_code": "AOYI-1", "attrs": {"Size": "one-size"},
         "sale_price": 13.25, "stock": 5},
    ]
    variations = {"Color": ["Green", "Pink", "White"]}
    kept, vars_, unknown = _filter_variants_by_declaration(
        "Color:Blue", skus, variations
    )
    assert kept == [], "Blue isn't in variations → should reject"
    assert "Color:Blue" in unknown


if __name__ == "__main__":
    test_empty_declaration_returns_everything()
    test_two_groups_cartesian()
    test_single_group_flat_allowlist()
    test_case_and_whitespace_insensitive()
    test_unknown_value_but_partial_match()
    test_zero_match_returns_empty_and_flags()
    test_flat_variations_key_not_in_sku_attrs()
    test_multi_group_seen_leak_reports_unreachable_value()
    test_attribute_prefix_stripped()
    test_full_width_colon_prefix_stripped()
    test_token_match_substring_pink_matches_1pc_pink()
    test_token_match_not_pinkeye()
    test_single_value_attribute_enriched_to_all_skus()
    test_multi_value_top_level_attr_uses_user_hint()
    test_multi_value_top_level_attr_rejects_when_user_hint_absent()
    print("ALL PASS")
