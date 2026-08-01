# AU Pipeline Refactor — Single-Template Schema, AI-Key Packaging, Parallelism

**Date:** 2026-08-01
**Scope:** `shein-extract-au` (澳洲站)
**Status:** Approved for implementation planning.

## Motivation

Five improvements collected in one release. Item 5 is the structural change that motivates the others; items 1/3/4 come along because they interact with (or are unblocked by) the new schema.

| # | Complaint | Root cause | This spec |
|---|---|---|---|
| 1 | eBay titles look dumb (raw first-80 chars + `NEW ` prefix) | AI path silently fails on employee installs — `-au` build has **no** `key_store.py` (only `.env`, which employee installers lack) → fallback to `_make_ebay_title()` rule | Port US project's `make_key_store.py` + `key_store.py` bundling into `-au` build |
| 2 | Variant scan not always accurate | Unconfirmed; suspected per-variant price mismatch | **Not fixed this release.** Item 5 lets the user pre-declare variants, sidestepping most of the pain. Record in `docs/known-issues.md`. |
| 3 | Random 4–12 s + long pause every 18 URLs looks robotic without real benefit | Historical anti-rate-limit hedging | Replace with fixed **3 s** between batches |
| 4 | Would like to run 5–10 tabs at once | Current loop is serial | **3 concurrent tabs** via `ThreadPoolExecutor` (conservative pick) |
| 5 | Two output files (batch xlsx + txt in `图片-<sku>/`) drift out of sync; input xlsx is a separate schema | Historical growth — no single source of truth | New Chinese-header template becomes input + primary output. Batch xlsx retained as backup, English headers unchanged. |

## §1 — New Template Schema (item 5)

**Source of truth:** the new template `C:\Users\ak\Downloads\Template - Test.xlsx` (sheet name = store, e.g. `ZR1`).

**Columns A–L (Chinese headers):**

| Col | Header | Direction | Semantics |
|-----|--------|-----------|-----------|
| A | 编号 | user-fill | Sequence number; also the media folder name (identical role to old `Seq`) |
| B | 链接 | user-fill | Shein product URL |
| C | 原价 | user-fill | Manual sale price (AUD). Blank → use scraped `data['price']` |
| D | 运费 | user-fill | Manual shipping (AUD). Blank → use scraper-derived shipping |
| E | 变体 | user-fill | Blank = scrape **all** variants. Filled = restrict scrape/output to the listed values. Free-form text; recommended format `Black, Red / M, L` (color / size, comma-separated within each group) |
| F | 日期 | write | Run date `YYYY-MM-DD` |
| G | 状态 | write | `Done` / `Failed` / `Delisted` |
| H | 希音价格 | write | Scraped sale price. If multi-variant with differing prices → range string, e.g. `$9.99–$14.99`. Purely a sanity-check against column C — never fed into any formula. |
| I | 希音标题 | write | Raw scraped `h1` title (unmodified) |
| J | eBay 标题 | write | AI-generated (Claude Haiku 4.5) |
| K | eBay 价格 | write | `(C or scraped-price) × 2 + (D or scraped-shipping)` — single scalar |
| L | 库存 | write | One line, variants slash-separated: `Black-M: 12 / Black-L: 缺货 / Red-M: 少货 3` |

**Behavioral rules:**

- **Trigger:** unchanged — `run_excel.py` scans `SUBMITTED_DIR` and processes every top-level `.xlsx` (or the file given as CLI arg).
- **Row selection:** process rows where `B (链接)` is filled AND `F (日期) + G (状态)` are both empty.
- **Re-run:** clear both `F` and `G` to re-pick a row (unchanged from current behavior).
- **eBay-title generation:** AI-first (Haiku), fallback to `_make_ebay_title()` rule. Same title lands in both `J` column and the per-SKU `eBay上架描述.txt`.
- **Variant filter (column E) — grammar and matching:**
  - **Grammar:** groups separated by `/`, values within a group comma-separated. Example: `"Black, Red / M, L"` = 2 groups × 2 values = 4 combos (Black-M, Black-L, Red-M, Red-L). Group order is not significant; matching is whitespace-insensitive and case-insensitive.
  - **Single group (no `/`):** treated as a flat allow-list matched against any attribute value. Example: `"Black, Red"` on a color+size product keeps every size for Black and Red.
  - **Matching against `sku_prices`:** a variant is kept only if **every non-empty group** in E has at least one value equal (case/whitespace-insensitive) to one of that SKU's attribute values.
  - **No match:** if E is filled but zero scanned variants match → treat product as failed with status `Failed` and reason `variant_filter_no_match` (surfaced in debug log); no partial write.
  - **Unmatched-only-partial:** if some E values don't exist on the page but at least one does → keep matches, log warning listing unknown values.

## §2 — Old Format Removal

The old English-header input format (`A=Seq / B=Website / C=Price / D=Date / E=Status / H=Web price`) is **removed** from `run_excel.py`. Any file whose header row does not include `链接` in column B is skipped with a warning. Downstream call-sites in `run_excel.py` that referenced English column indices are updated to Chinese ones or removed.

The **batch summary xlsx** (`{store}-{minSeq}-{maxSeq}-{YYYYMMDD}.xlsx` under `OUTPUT_ROOT_2ND/{store}/`) is unchanged — same 14-column English-header layout, same variant sub-row expansion, same embedded pictures. Retained as a backup / audit view.

The per-SKU folder `图片-<sku>/` and its `eBay上架描述.txt` are unchanged in content and format.

## §3 — AI Key Packaging (item 1)

Mirror the US project's build-time key bundling:

- Copy `make_key_store.py` (or equivalent) from `C:\Users\ak\Desktop\Claude\shein-extract\` to `-au`.
- Add its invocation to `build.bat` before `pyinstaller` runs, so `key_store.py` (containing the obfuscated `ANTHROPIC_API_KEY`) is included in the compiled `dist/`.
- `key_store.py` remains gitignored; `.env` fallback stays for source-run.
- Verify by checking `_get_api_key()` returns non-empty inside a compiled build.

## §4 — Fixed 3-Second Pacing (item 3)

- Delete constants: `INTER_URL_DELAY_MIN`, `INTER_URL_DELAY_MAX`, `LONG_PAUSE_EVERY`, `LONG_PAUSE_MIN`, `LONG_PAUSE_MAX`.
- Rewrite `_inter_url_pause(i, total)` to `time.sleep(3)` if `i < total`.
- Under parallelism (§5), the 3 s becomes the gap **between batches of 3**, not between individual URLs.

## §5 — 3-Tab Parallelism (item 4)

**Architecture:**

- Wrap the current per-URL scrape body into a function `scrape_url(url, seq, price, port) -> record`.
- Drive it with `ThreadPoolExecutor(max_workers=3)` batched over the URL list.
- Each worker owns its own tab (`_new_tab` → `tab_id` → `ws_url`) and closes it in `finally`.
- CDP is thread-safe at the WebSocket level (each `_cdp_once` opens a fresh WS), so no shared-state contention on the Chrome side.

**Concurrency safety:**

- **Excel writes** in `run_excel.py` (`ws.cell(row, col).value = …`) are serialized with a `threading.Lock`. Same for `safe_save`.
- **Image download pool:** reduce `IMAGE_DOWNLOAD_WORKERS` from **8 → 4** (3 concurrent products × 4 workers = 12 sockets, comfortable).
- **Rate-limit counter:** the `RATE_LIMIT_CONSECUTIVE=3` guard becomes a `threading.Lock`-protected shared counter. On trip → `Executor.shutdown(cancel_futures=True)` and raise `RateLimitError`.
- **Session warmup** (`_ensure_shein_session`) still runs **once**, before the executor is created.
- **Captcha / signin block** detection stays per-tab (unchanged logic). A tab blocked by captcha waits (existing 300 s poll) without blocking siblings — but the executor's max workers cap means one stuck tab reduces effective parallelism by 1 until it clears.
- **Batch pacing:** after each future in a batch of 3 resolves, the outer loop waits 3 s before submitting the next batch (item 3 rule).

**Not changing:**

- CDP port stays `9223` (single Chrome instance, multiple tabs).
- Persistent profile dir stays `~/shein-cdp-profile-au` (shared across tabs — a single browser profile is correct here).

## §6 — Item 2 Deferred

Add a top-level `docs/known-issues.md` with a single entry:

> **Variants**: per-variant price/stock mapping (`sku_prices`) may occasionally mis-associate with attribute values. Symptoms: row L shows a `-` where a stock number is expected, or `sku_prices` in the txt shows a variant not actually offered on the page. Item 5 lets the user pre-declare the variant list in column E, which sidesteps most cases. To fix root cause, capture the failing URL and the raw `window.gbRawData.modules.saleAttr` JSON before proposing a patch.

## §7 — Testing

- **Unit tests:**
  - `test_template_reader.py`: given a fixture template with mixed filled/empty rows, verify only ready rows are picked.
  - `test_variant_filter.py`: given a scraped product with 8 variants and E column `"Black, Red / M"`, verify only Black-M and Red-M appear in output.
  - `test_price_calc.py`: verify eBay-price formula honors C/D overrides and falls back to scraped values.
  - `test_stock_summary.py`: verify L column format (`Black-M: 12 / Black-L: 缺货 / Red-M: 少货 3`).
  - Extend `test_variant_merge.py` (existing) to cover new template columns.

- **Manual verification:** run one sheet end-to-end after implementation with:
  - 1 single-variant product (Done, no ranges)
  - 1 multi-variant product with filled E column (only listed variants in L)
  - 1 multi-variant product with blank E column (full L)
  - 1 delisted product (Delisted status)
  - 1 URL that fails 3 times (Failed status)

- **AI-key bundling check:** compile via `build.bat`, launch installed exe, run a single URL, confirm the eBay title is NOT the fallback (fallback always starts with `NEW ` for anything ≤ 80 chars — regressions will be visible).

## §8 — Migration & Compatibility

- **Existing employees** need to redistribute the exe (item 1's key bundling ships as part of the same installer bump).
- **Existing in-flight input xlsx** (English-header) will stop being processed. Communicate before rollout; provide a one-off conversion snippet if needed:
  ```python
  # convert old → new headers (columns rearrange too — see design §1)
  # Not part of scope; user handles manually.
  ```
- **Batch xlsx consumers** (if any) are unaffected — English format preserved.

## §9 — Out of Scope

- Prompt improvements to `_EBAY_TITLE_PROMPT` beyond current AU version.
- Per-variant eBay prices in column K (single scalar only — user commits to filling C and D per row).
- Cross-batch/cross-sheet aggregation.
- Variant accuracy fixes (item 2).
- Any change to the `图片-<sku>` folder layout or txt file content.

## §10 — Success Criteria

- New template flows: user fills A/B/C/D (+ optional E), runs `run_excel.py`, sees F–L filled correctly and the batch xlsx + per-SKU folders produced as before.
- Employee installer: eBay titles reflect AI output, not the rule fallback.
- Runtime: on a batch of 15 URLs, wall-clock is roughly 1/3 of the current serial time (rough acceptance — captcha/retries dominate outliers).
- Existing tests still pass; new tests listed in §7 pass.
