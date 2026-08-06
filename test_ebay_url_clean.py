"""Tests for _clean_ebay_url — strips tracking / query params from
eBay listing URLs, keeps a clean sharable form."""
from ebay_scraper import _clean_ebay_url


def test_strip_campaign_param():
    u = "https://www.ebay.com.au/itm/123456789?campid=5338&customid=xyz"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/123456789"


def test_strip_hash_param():
    u = "https://www.ebay.com.au/itm/12345?hash=item1abc:g:XYZAA"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/12345"


def test_strip_trkparms_and_all_others():
    u = ("https://www.ebay.com.au/itm/12345"
         "?_trkparms=abc%3Dxyz&mkevt=1&mkcid=1&mkrid=705-53470-19255-0"
         "&campid=5338&customid=default")
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/12345"


def test_preserve_slug_variant():
    # eBay sometimes uses /itm/<slug>/<id> — keep both.
    u = "https://www.ebay.com.au/itm/Cute-Ceramic-Mug/22334455?campid=5338"
    assert _clean_ebay_url(u) == "https://www.ebay.com.au/itm/Cute-Ceramic-Mug/22334455"


def test_already_clean_passthrough():
    u = "https://www.ebay.com.au/itm/98765"
    assert _clean_ebay_url(u) == u


def test_non_ebay_passthrough():
    # Defensive: if extractor ever picks up a non-eBay href, don't mangle it.
    u = "https://example.com/some?path=1&other=2"
    assert _clean_ebay_url(u) == u


if __name__ == "__main__":
    test_strip_campaign_param()
    test_strip_hash_param()
    test_strip_trkparms_and_all_others()
    test_preserve_slug_variant()
    test_already_clean_passthrough()
    test_non_ebay_passthrough()
    print("ALL PASS")
