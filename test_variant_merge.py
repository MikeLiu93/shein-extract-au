"""Test: all colors from mainSaleAttribute must land in `variations`, even when
the page pre-selected one color via ?main_attr=... .

Reproduces the bug from product p-454706293 (2 colors, only the selected one
was detected). Run: python test_variant_merge.py
"""
from shein_scraper import _merge_main_sale_attr_colors, _make_ebay_title


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


if __name__ == "__main__":
    test_preselected_color_keeps_all_colors()
    test_no_main_sale_attrs_leaves_variations_unchanged()
    test_existing_key_casing_preserved_and_deduped()
    test_single_color_unchanged()
    test_title_omits_color_when_multiple()
    test_title_keeps_color_when_single()
    print("ALL PASS")
