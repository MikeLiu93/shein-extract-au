# Known Issues (deferred)

## Variants — per-variant price/stock may mis-associate

**Symptom:** in the batch xlsx or in `图片-<sku>/eBay上架描述.txt`, a variant
occasionally shows a wrong price, an empty stock, or a variant that doesn't
actually exist on the page. In the new template's `库存` column, the same
mismatch surfaces as a `-` where a number is expected.

**Suspected cause:** the mapping between `sku_prices` entries and their
attribute values inside `window.gbRawData.modules.saleAttr.multiLevelSaleAttribute.sku_list`
is fragile on some product layouts; the extraction JS in `shein_scraper.py::_JS`
sometimes hands the wrong attr row to the wrong SKU.

**Workaround (in place since 2026-08-01):** the new template's `E 变体` column
lets the operator pre-declare exactly which variants to list. When filled,
non-matching `sku_prices` entries are dropped before write, so mis-associations
outside the declared set do not appear.

**When to actually fix:** capture the failing URL + full `window.gbRawData`
JSON, then refine the `extractSkuPrices()` / attribute-mapping logic. Do not
attempt a speculative fix without a reproduction.
