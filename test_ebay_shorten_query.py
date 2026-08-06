"""Tests for _shorten_query — turns the 80-char AI eBay title (which returns
0 results on eBay AU search) into a 3-5 word query that finds relevant
listings. Strips leading quantity tokens ('12pcs', '100Pcs', '1/2pcs') and
caps at ~5 content words / ~50 chars."""
from ebay_scraper import _shorten_query


def test_strip_leading_pcs_token():
    # '12pcs' at the front is a Chinese-listing tell; strip it so search matches
    # nouns instead of the specific pack size.
    q = _shorten_query("12pcs Glitter Pumpkin Cupcake Toppers Halloween Thanksgiving Party Cake")
    assert not q.lower().startswith("12pcs"), q
    assert "Glitter" in q and "Pumpkin" in q


def test_strip_hundred_pcs_variant():
    q = _shorten_query("100pcs Bamboo Cocktail Picks 4.7\" Animal Shaped Party Cake")
    assert not q.lower().startswith("100pcs"), q
    assert "Bamboo" in q


def test_strip_slash_pcs_variant():
    q = _shorten_query("1/2pcs Creative Hand-Painted Cute Bear Ceramic Mug With Bow Tie")
    assert not q.lower().startswith("1/2pcs"), q
    assert "Creative" in q


def test_no_leading_quantity_passthrough():
    q = _shorten_query("Vintage French Cream Floral Embossed Tea Cup")
    assert q.startswith("Vintage"), q


def test_cap_at_five_words():
    q = _shorten_query("Alpha Bravo Charlie Delta Echo Foxtrot Golf Hotel India")
    # 5 words max
    assert len(q.split()) <= 5, q


def test_cap_at_fifty_chars():
    q = _shorten_query("Verylongword " * 10)
    assert len(q) <= 50, q


def test_empty_returns_empty():
    assert _shorten_query("") == ""
    assert _shorten_query(None) == ""


def test_pack_variant_also_stripped():
    q = _shorten_query("1 Set Adorable Ceramic Mug With Lid Cute 3D Cloud")
    # "1 Set" is 2 tokens both quantity-y; strip both.
    assert not q.lower().startswith("1"), q
    assert "Ceramic" in q or "Adorable" in q


if __name__ == "__main__":
    test_strip_leading_pcs_token()
    test_strip_hundred_pcs_variant()
    test_strip_slash_pcs_variant()
    test_no_leading_quantity_passthrough()
    test_cap_at_five_words()
    test_cap_at_fifty_chars()
    test_empty_returns_empty()
    test_pack_variant_also_stripped()
    print("ALL PASS")
