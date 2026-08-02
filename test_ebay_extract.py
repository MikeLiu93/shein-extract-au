"""Tests for _extract_from_page — takes the JS return (list of dicts with
raw strings from the eBay search page) and yields a normalized list of
EbayHit instances, cleaning URLs and computing delivered_price."""
from ebay_scraper import _extract_from_page, EbayHit


def _raw(title="Foo Widget", price="AU $12.99", postage="+ AU $6.95 postage",
         url="https://www.ebay.com.au/itm/111?campid=5338&hash=x",
         img="https://i.ebayimg.com/thumbs/foo.jpg"):
    return {"title": title, "price": price, "postage": postage,
            "url": url, "image_url": img}


def test_two_full_hits():
    raw = [_raw(title="A"), _raw(title="B", price="AU $15.00", postage="Free postage",
                                 url="https://www.ebay.com.au/itm/222?campid=1")]
    hits = _extract_from_page(raw)
    assert len(hits) == 2
    assert hits[0].title == "A"
    assert hits[0].sticker_price == 12.99
    assert hits[0].postage == 6.95
    assert hits[0].delivered_price == 19.94
    assert hits[0].url == "https://www.ebay.com.au/itm/111"
    assert hits[1].sticker_price == 15.00
    assert hits[1].postage == 0.0
    assert hits[1].delivered_price == 15.00


def test_postage_unparseable_falls_back_to_sticker():
    raw = [_raw(postage="Contact seller for postage")]
    hits = _extract_from_page(raw)
    assert len(hits) == 1
    assert hits[0].postage is None
    assert hits[0].delivered_price == hits[0].sticker_price  # = 12.99


def test_price_unparseable_skips_that_hit():
    # No usable price → drop this hit (can't rank without a price).
    raw = [_raw(price="SEE PRICE IN CART"), _raw(title="Real Hit")]
    hits = _extract_from_page(raw)
    assert len(hits) == 1
    assert hits[0].title == "Real Hit"


def test_empty_input_returns_empty_list():
    assert _extract_from_page([]) == []


def test_at_most_two_returned_even_if_more_input():
    raw = [_raw(title=f"Hit{i}") for i in range(5)]
    hits = _extract_from_page(raw)
    assert len(hits) == 2
    assert hits[0].title == "Hit0"
    assert hits[1].title == "Hit1"


def test_url_is_cleaned():
    raw = [_raw(url="https://www.ebay.com.au/itm/Nice-Slug/333?campid=5338&_trkparms=x")]
    hits = _extract_from_page(raw)
    assert hits[0].url == "https://www.ebay.com.au/itm/Nice-Slug/333"


if __name__ == "__main__":
    test_two_full_hits()
    test_postage_unparseable_falls_back_to_sticker()
    test_price_unparseable_skips_that_hit()
    test_empty_input_returns_empty_list()
    test_at_most_two_returned_even_if_more_input()
    test_url_is_cleaned()
    print("ALL PASS")
