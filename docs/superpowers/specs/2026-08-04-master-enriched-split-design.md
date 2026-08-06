# Master-Enriched Split — Design

**Date:** 2026-08-04
**Scope:** `shein-extract-au` (澳洲站)
**Status:** Approved for implementation planning.
**Follow-on to:** `2026-08-01-au-template-refactor-design.md` (single-template schema) and `2026-08-02-ebay-price-check-design.md` (eBay price check).

## Motivation

Since the 2026-08-01 refactor, the pipeline uses a **single Chinese-header template** as both input (user fills links + costs) and output (scripts write scrape results back to the same file). A recent file-corruption incident put the operator's entire link inventory at risk — the scripts and Excel share write access to one file, so any corruption or partial write is catastrophic.

This spec splits the single file into **three layers**:

1. **Master (输入)** — the operator's link inventory. Scripts NEVER write to it. Only the operator edits it. Six columns.
2. **Enriched (输出)** — an append-only cumulative view. Scripts append one new row per (product × run). Preserves full history: a seq processed 5 times appears as 5 rows dated differently. Twenty-one columns matching the current template.
3. **Batch xlsx** — the existing per-run summary (`{store}-{minSeq}-{maxSeq}-{YYYYMMDD}.xlsx`). Unchanged.

The trade-off: two files instead of one for the operator to look at, plus enriched grows over time. The safety win is worth it — a corrupt enriched file means "lost the append history since last backup"; a corrupt master (impossible under this design) meant "lost all link inventory."

## §1 — Master (输入) Schema

Six columns. This IS the source of truth for links + operator's costing decisions.

| Col | Header | Type | Semantics |
|---|---|---|---|
| A | 编号 | int | Sequence number; also the media folder name |
| B | 链接 | str | Shein product URL |
| C | 原价 | float | Manual sale price (AUD). Blank → use scraped price. |
| D | 运费 | float | Manual shipping (AUD). Blank → use scraped shipping. |
| E | 变体 | str | Variant filter declaration (per existing rules). Blank → all. |
| F | 是否要跑 | str | **Case-insensitive `"Y"` (whitespace-stripped) triggers this row.** Any other value (`""`, `"N"`, `"完成"`, ...) → skip. |

Rules:
- Script **never writes** to the master file (no `wb.save()` on the master path).
- Sheet name = store name (e.g. `ZR1`). Same convention as before.
- Non-template sheets (B1 ≠ `"链接"`) are silently skipped.
- Operator opens the master in Excel freely; script only reads (no lock conflict with Excel's read-only mode).

## §2 — Enriched (输出) Schema

Twenty-one columns — **identical to the current template layout** (2026-08-01 spec §1 + 2026-08-02 spec §1). No new columns.

```
A 编号 | B 链接 | C 原价 | D 运费 | E 变体          ← copied from master at scrape time
F 日期 | G 状态 | H 图片 | I 希音价格 | J 希音标题     ← run_excel writes
K eBay标题 | L eBay价格 | M 库存
N eBay搜索日期 | O eBay同类低价 | P 低价链接
Q eBay同类高价 | R 高价链接                       ← ebay_price_check writes
S Shein重跑日期 | T 更新价格 | U 更新库存           ← reserved for future
```

Rules:
- Every write is an **append** — never in-place edit of existing rows (except the eBay N-R block, which fills row-in-place by seq+日期 lookup — see §4).
- Sheet layout mirrors master: one sheet per store (`ZR1`, ...).
- Enriched file created on first run if absent (with header row).
- Enriched **sheet** auto-created (with header row) if the file exists but the store sheet is missing — this handles onboarding a new store without asking the operator to seed anything.
- If enriched exists AND the store sheet exists but headers don't match the 21-column layout exactly → hard error, script stops (don't pollute stale schema).
- Excel formatting (row height, image anchors) applied per row on append.

## §3 — Trigger & Row Selection

**`run_excel.py` (Shein scrape):**
- Read every sheet in master. For each row where `F (是否要跑)` normalizes to `"Y"`, process it.
- After processing (success or fail), append one row to enriched's same-named sheet.
- The `F=Y` value stays on the master (script does not clear). Operator decides when to un-Y.
- If operator leaves `F=Y` and re-runs later, another enriched row appended — this is the intended history mode.

**`ebay_price_check.py`:**
- Read enriched. For each row where `G=Done` AND `N (eBay 搜索日期)` is empty AND `K (eBay标题)` non-empty → search eBay and fill `N-R`.
- Ebay check writes IN-PLACE to enriched (same row that Shein just appended). This is the one exception to "never edit existing rows" — the row was appended empty in N-R, eBay check fills them.
- The N-R fill is idempotent by design (empty N triggers, filled N skips), so re-runs don't over-write.
- `ebay_price_check.py` runs INDEPENDENTLY of `run_excel.py` — no notion of "current run". It scans every enriched row across every sheet, filters on `G=Done AND N empty AND K non-empty`, updates in-place by openpyxl row index. No cross-file lookup or seq-matching needed at this layer.

## §4 — Row Identity & Append Semantics

**Master → Enriched mapping:**
- When Shein processes master row seq=5, it copies A-E verbatim to a new enriched row, plus fills F-M with scrape results.
- The enriched row has no direct "back-reference" to the master row (no UUID). Its identity in enriched is `(seq, F 日期)`.
- If master row seq=5's URL changes and F stays `Y`, next run appends a new enriched row with the new URL. The old enriched row is untouched.

**Enriched integrity:**
- Rows are ordered by append time (no sort applied after write).
- Header row (row 1) never re-written after initial creation.
- The eBay in-place update in §3 finds the target row by seeking within a specific batch (the rows just appended by the current run) — see implementation contract below.

## §5 — Implementation Contract for the Append Model

**`run_excel.process_excel(master_path, enriched_path)` shape:**
1. Load master (read-only in intent — do not call `wb.save()` on it).
2. For each sheet:
    - Collect pending rows (F=Y).
    - Load or create enriched sheet with same name.
    - For each pending row: scrape → build enriched row dict → append at `ws.max_row + 1` under enriched sheet.
    - After all pending rows in this sheet: `safe_save(enriched)`.
3. Batch xlsx generation happens as before, using the same records list.

**`ebay_price_check.process_excel(enriched_path)` shape:**
1. Load enriched (read + write).
2. For each sheet:
    - Collect rows where `G=Done AND N=empty AND K non-empty`.
    - For each: search eBay → write N-R via `ws.cell(row, ...).value = ...` (in-place row update).
3. `safe_save(enriched)`.

## §6 — Config Changes

**`.env`** (per-machine):
```
SHEIN_SUBMITTED_DIR=<dir containing both files>
SHEIN_INPUT_FILENAME=澳洲希音链接 (输入) - ZR.xlsx    ← master
SHEIN_OUTPUT_FILENAME=澳洲希音链接 (输出) - ZR.xlsx   ← enriched (NEW)
SHEIN_OUTPUT_DIR=<where per-batch xlsx + 图片-<sku>/ go>
```

**`config.py`**:
- Add `OUTPUT_FILENAME = os.environ.get("SHEIN_OUTPUT_FILENAME", "").strip()`.
- No default value. Scripts that require it (`run_excel.py`, `ebay_price_check.py`) call a small `_require_output_filename()` helper that raises `SystemExit` with the clear message "SHEIN_OUTPUT_FILENAME not set in .env / config.env; run setup_wizard or edit the file manually." Failure is at first use, not at config load — keeps `config.py` importable in test contexts.

**`setup_wizard.py`**:
- Rename existing "指定输入文件名（可选）" field to **"输入表文件名（必填）"** — now required.
- Add new field **"输出表文件名（必填）"** with corresponding validation.
- Both write to the config.env alongside the existing keys.

## §7 — File Structure & File Ownership

**Locations (all in `SHEIN_SUBMITTED_DIR`):**
- Master: `SUBMITTED_DIR / <input filename>`
- Enriched: `SUBMITTED_DIR / <output filename>`
- Batch xlsx + `图片-<sku>/`: `SHEIN_OUTPUT_DIR / <store> /` (unchanged)

**Write ownership:**
- Master → only operator
- Enriched → only scripts (`run_excel.py` appends; `ebay_price_check.py` fills N-R in-place)
- Batch xlsx → only `run_excel.py` (per-run new file)
- Media folders → only `run_excel.py` per SKU

## §8 — Migration

Operator handles manually (already done):
- Existing template renamed to `澳洲希音链接 (输入) - ZR.xlsx`, trimmed to 6 columns, F column filled with "Y" for rows to process.
- Empty `澳洲希音链接 (输出) - ZR.xlsx` prepared with the 21-column header row (or auto-created by script on first run).

Historical scraped data (from prior 21-column runs) NOT auto-migrated. The batch xlsx files serve as backup for pre-refactor snapshots.

## §9 — Failure Modes

| Situation | Handling |
|---|---|
| Master file missing | Fail-fast at startup; clear error message |
| Master schema mismatch (`B1 ≠ "链接"`) | Sheet skipped with log warning |
| Enriched file missing | Auto-create with 21-col header row |
| Enriched schema mismatch (any header differs) | Fail-fast; do not touch the file |
| Enriched file locked by Excel | `safe_save` writes `- 输出2 - ZR.xlsx` (existing pattern) |
| Master row has F=Y but no URL | Skip with log warning |
| Master row's 编号 non-numeric | Skip with log warning (existing behaviour) |
| Script crash mid-append | Master unaffected (never opened for write); enriched may miss the crashed row → re-run picks it up (F still Y) |
| Two script runs in parallel | Not supported — operator's responsibility to avoid |

## §10 — Testing

**Unit** (plain-`assert`, follow project convention):
- `test_master_reader.py` — read a fresh 6-col master; skip non-Y rows; whitespace-strip on F.
- `test_enriched_appender.py` — append a row to a fresh workbook; create sheet with headers if absent; reject mismatched headers.
- `test_ebay_inplace_update.py` — update N-R on a specific enriched row (matched by seq + 日期) without disturbing other rows.

**Manual smoke:** end-to-end run against the real ZR files after refactor.

## §11 — Out of Scope

- Detecting master edits (URL change, price change) — treated as append (new snapshot in enriched).
- Dedup / cleanup of enriched (growing table). Operator manages Excel filtering.
- Rollback of enriched writes on partial failure.
- Concurrent script runs.
- Master history / audit trail (git-tracked or otherwise).
- UI for browsing enriched history.
- Compressing multi-run enriched into a "latest snapshot" view (achievable via Excel filter on latest 日期 per seq).

## §12 — Backwards Incompatibility

- Old single-file template no longer processed. Operator's `澳洲希音链接汇总 - ZR.xlsx` (21-col mixed) will be silently ignored (B1 header may still be `链接`, but no F=Y trigger will fire since old F was 日期, not 是否要跑).
- Existing employee installs (v0.2.1) will need a wizard re-run after upgrading to pick up the new output-filename field. Old configs missing `SHEIN_OUTPUT_FILENAME` will fail with the clear error message from §6.
- Version bump to 0.3.0 (minor bump — new required config).
