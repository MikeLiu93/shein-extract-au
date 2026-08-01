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


if __name__ == "__main__":
    test_empty_declaration_returns_everything()
    test_two_groups_cartesian()
    test_single_group_flat_allowlist()
    test_case_and_whitespace_insensitive()
    test_unknown_value_but_partial_match()
    test_zero_match_returns_empty_and_flags()
    print("ALL PASS")
