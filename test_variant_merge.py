"""Test: mainSaleAttribute → variations merge.

Two modes:
  1. URL passed → use the URL's goods_id to pick THIS URL's canonical color
     (2026-09 fix — was: overwrote with all SPU colors, losing "which color am I?").
  2. URL absent / not matched → collect the full SPU color list (old behavior,
     needed so DOM-lost-colors case p-454706293 doesn't regress).

Run: python test_variant_merge.py
"""
from shein_scraper import (
    _merge_main_sale_attr_colors, _make_ebay_title,
    _pick_current_color_from_main_sale_attrs,
)


def test_preselected_color_keeps_all_colors():
    # Captured live from p-454706293?main_attr=27_279167171:
    # DOM gave only the selected color; main_sale_attrs has both.
    variations = {"Color": ["1PC Walnut Color"]}
    main_sale_attrs = [
        {"attr_value_name": "1PC Natural Wood Color", "attr_name": "Color",
         "goods_sn": "sr260425153678508295008", "is_current": False},
        {"attr_value_name": "1PC Walnut Color", "attr_name": "Color",
         "goods_sn": "sr260425153678508216698", "is_current": False},
    ]
    merged = _merge_main_sale_attr_colors(variations, main_sale_attrs)
    assert merged["Color"] == ["1PC Natural Wood Color", "1PC Walnut Color"], merged


def test_no_main_sale_attrs_leaves_variations_unchanged():
    variations = {"US Size": ["US 8", "US 9"]}
    assert _merge_main_sale_attr_colors(variations, []) == {"US Size": ["US 8", "US 9"]}


def test_existing_key_casing_preserved_and_deduped():
    variations = {"color": ["A"]}  # lowercase key from DOM
    msa = [{"attr_value_name": "A", "attr_name": "Color"},
           {"attr_value_name": "A", "attr_name": "Color"},  # dup
           {"attr_value_name": "B", "attr_name": "Color"}]
    merged = _merge_main_sale_attr_colors(variations, msa)
    assert merged["color"] == ["A", "B"], merged  # keep DOM key casing, dedupe


def test_single_color_unchanged():
    variations = {"Color": ["Only Red"]}
    msa = [{"attr_value_name": "Only Red", "attr_name": "Color"}]
    assert _merge_main_sale_attr_colors(variations, msa)["Color"] == ["Only Red"]


def test_title_omits_color_when_multiple():
    title = "Foldable Bamboo Grain Armrest Tray, Space-Saving Design"
    multi = {"Color": ["1PC Natural Wood Color", "1PC Walnut Color"]}
    out = _make_ebay_title(title, multi)
    assert "Natural Wood" not in out and "Walnut" not in out, out


def test_title_keeps_color_when_single():
    title = "Foldable Bamboo Grain Armrest Tray, Space-Saving Design"
    single = {"Color": ["1PC Walnut Color"]}
    out = _make_ebay_title(title, single)
    assert "Walnut" in out, out


# ── 2026-09: URL-anchored current color ────────────────────────────────────

# Real product data captured from au.shein.com/SHEIN-EZwear-...-p-73768996.html
# The URL slug is "Beige Color... Butter Yellow Dress" but the internal
# color for goods_id 73768996 is 'Apricot'. That mismatch was the reason
# employees typing "Color:Beige" or "Color:Yellow" got zero results.
EZWEAR_URL = ("https://au.shein.com/SHEIN-EZwear-Women-s-Casual-Long-Dress"
              "-In-Beige-Color-Summer-Butter-Yellow-Dress-p-73768996.html")
EZWEAR_MSA = [
    {"attr_value_name": "Apricot",  "attr_name": "Color", "goods_id": "73768996"},
    {"attr_value_name": "Apricot",  "attr_name": "Color", "goods_id": "374146635"},
    {"attr_value_name": "Black",    "attr_name": "Color", "goods_id": "424803192"},
    {"attr_value_name": "White",    "attr_name": "Color", "goods_id": "60251625"},
]


def test_pick_current_color_by_url_goods_id():
    got = _pick_current_color_from_main_sale_attrs(EZWEAR_URL, EZWEAR_MSA)
    assert got == {"attr_name": "Color", "value": "Apricot"}, got


def test_pick_current_color_returns_none_when_url_missing():
    assert _pick_current_color_from_main_sale_attrs(None, EZWEAR_MSA) is None
    assert _pick_current_color_from_main_sale_attrs("", EZWEAR_MSA) is None


def test_pick_current_color_returns_none_when_url_has_no_goods_id():
    assert _pick_current_color_from_main_sale_attrs(
        "https://au.shein.com/some-page.html", EZWEAR_MSA
    ) is None


def test_pick_current_color_returns_none_when_url_gid_not_in_msa():
    assert _pick_current_color_from_main_sale_attrs(
        "https://au.shein.com/whatever-p-99999999.html", EZWEAR_MSA
    ) is None


def test_merge_with_url_writes_only_the_url_color():
    """The real bug: DOM might guess wrong (or right), but the URL goods_id is
    the ground truth. Merge must collapse variations['Color'] to just Apricot,
    NOT the full 4-color SPU list."""
    variations = {"Color": ["Butter Yellow"], "Size": ["XS", "S", "M"]}
    merged = _merge_main_sale_attr_colors(variations, EZWEAR_MSA, url=EZWEAR_URL)
    assert merged["Color"] == ["Apricot"], merged
    assert merged["Size"] == ["XS", "S", "M"], "unrelated dims must not change"


def test_merge_with_url_creates_color_key_when_absent():
    """DOM missed Color entirely — URL still anchors it."""
    variations = {"Size": ["XS", "S", "M"]}
    merged = _merge_main_sale_attr_colors(variations, EZWEAR_MSA, url=EZWEAR_URL)
    assert merged["Color"] == ["Apricot"], merged


def test_merge_with_url_preserves_existing_key_casing():
    """DOM used lowercase 'color'; merge must NOT introduce a duplicate uppercase key."""
    variations = {"color": ["something"]}
    merged = _merge_main_sale_attr_colors(variations, EZWEAR_MSA, url=EZWEAR_URL)
    assert "color" in merged and "Color" not in merged, merged
    assert merged["color"] == ["Apricot"], merged


def test_merge_without_url_falls_back_to_full_spu_list():
    """No URL passed — regression guard for the p-454706293 case (DOM lost
    one color; merge must collect the full list so eBay title logic
    (multi-color → omit) still works)."""
    variations = {"Color": ["1PC Walnut Color"]}
    msa = [
        {"attr_value_name": "1PC Natural Wood Color", "attr_name": "Color"},
        {"attr_value_name": "1PC Walnut Color", "attr_name": "Color"},
    ]
    merged = _merge_main_sale_attr_colors(variations, msa)
    assert merged["Color"] == ["1PC Natural Wood Color", "1PC Walnut Color"], merged


def test_merge_with_url_but_no_goods_id_match_falls_back():
    """URL passed but its goods_id isn't in main_sale_attrs — fall back to full-list."""
    variations = {"Color": ["A"]}
    msa = [
        {"attr_value_name": "A", "attr_name": "Color", "goods_id": "111"},
        {"attr_value_name": "B", "attr_name": "Color", "goods_id": "222"},
    ]
    merged = _merge_main_sale_attr_colors(
        variations, msa, url="https://au.shein.com/x-p-999.html"
    )
    assert merged["Color"] == ["A", "B"], merged


if __name__ == "__main__":
    test_preselected_color_keeps_all_colors()
    test_no_main_sale_attrs_leaves_variations_unchanged()
    test_existing_key_casing_preserved_and_deduped()
    test_single_color_unchanged()
    test_title_omits_color_when_multiple()
    test_title_keeps_color_when_single()
    test_pick_current_color_by_url_goods_id()
    test_pick_current_color_returns_none_when_url_missing()
    test_pick_current_color_returns_none_when_url_has_no_goods_id()
    test_pick_current_color_returns_none_when_url_gid_not_in_msa()
    test_merge_with_url_writes_only_the_url_color()
    test_merge_with_url_creates_color_key_when_absent()
    test_merge_with_url_preserves_existing_key_casing()
    test_merge_without_url_falls_back_to_full_spu_list()
    test_merge_with_url_but_no_goods_id_match_falls_back()
    print("ALL PASS")
