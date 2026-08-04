"""
Excel-based pipeline (澳洲站): read pending URLs from .xlsx worksheets (one per
store), scrape them, write results to OUTPUT_ROOT/{store}/, and update Date +
Status columns in the source Excel.

Usage:
    python run_excel.py                          # 扫描 SUBMITTED_DIR 下所有 .xlsx
    python run_excel.py "path/to/file.xlsx"      # 指定文件

Columns (strict, Chinese-header template schema):
    A: 编号       — sequence number (= output folder name). READ.
    B: 链接       — Shein product URL. READ.
    C: 原价       — manual sale price (AUD). READ. Blank → use scraped price.
    D: 运费       — manual shipping (AUD). READ. Blank → use scraped shipping.
    E: 变体       — variant filter declaration. READ. Blank → scrape all.
    F: 日期       — run date. WRITE.
    G: 状态       — Done / Failed / Delisted. WRITE.
    H: 图片       — embedded first product image (scaled). WRITE.
    I: 希音价格   — scraped price. Single value → numeric; multi-variant with
                    differing prices → string range "9.99–14.99". WRITE.
    J: 希音标题   — raw scraped h1 title. WRITE.
    K: eBay 标题  — AI-generated (Haiku, fallback to rule). WRITE.
    L: eBay 价格  — (C or scraped) × 2 + (D or scraped). WRITE.
    M: 库存       — one-line variant stock summary. WRITE.

Only rows where 链接 is filled AND 日期/状态 are both empty are processed.
Non-template sheets (missing '链接' in B1) are skipped.
"""

import argparse
import logging
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from shein_scraper import scrape_shein, RateLimitError, _add_picture_to_cell
from config import SUBMITTED_DIR, OUTPUT_ROOT_2ND as OUTPUT_ROOT, INPUT_FILENAME

logger = logging.getLogger("run_excel")
DEBUG_LOG_DIR = Path(__file__).resolve().parent / "debug_logs"

# ── New Chinese-header template schema ────────────────────────────────────────
# Cols A–M. See docs/superpowers/specs/2026-08-01-au-template-refactor-design.md §1.
# 2026-08-02: column H "图片" added (embedded product image); prices+titles shifted +1.
COL_SEQ, COL_URL, COL_PRICE, COL_SHIPPING, COL_VARIANT_FILTER = 1, 2, 3, 4, 5
COL_DATE, COL_STATUS, COL_PICTURE = 6, 7, 8
COL_WEB_PRICE, COL_SHEIN_TITLE = 9, 10
COL_EBAY_TITLE, COL_EBAY_PRICE, COL_STOCK = 11, 12, 13

# Enriched sheet columns A~U. See spec 2026-08-04-master-enriched-split-design.md §2.
EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体",
                    "日期", "状态", "图片", "希音价格", "希音标题",
                    "eBay标题", "eBay价格", "库存",
                    "eBay搜索日期", "eBay同类低价", "低价链接",
                    "eBay同类高价", "高价链接",
                    "Shein重跑日期", "更新价格", "更新库存"]

# ── Master (输入) schema ─────────────────────────────────────────────────────
# 主表只有 6 列，脚本只读。See spec 2026-08-04-master-enriched-split-design.md §1.
MASTER_COL_TRIGGER = 6  # F 是否要跑 (Y/空/其他)

MASTER_EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体", "是否要跑"]


def _read_master_pending_rows(ws) -> list[dict]:
    """Return list of dicts for rows in the master where F (是否要跑) normalizes
    to 'Y' (case-insensitive, whitespace-stripped). Non-master sheets return [].
    Shape: {row, seq, url, price, shipping, variant_filter}. Rows with invalid
    seq or missing URL are logged and skipped."""
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        url = ws.cell(r, COL_URL).value
        trigger = str(ws.cell(r, MASTER_COL_TRIGGER).value or "").strip().upper()
        if trigger != "Y":
            continue
        if not url:
            logger.info("  row %d: skip (F=Y but 链接 empty)", r)
            continue
        try:
            seq_int = int(seq) if seq is not None else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (编号 '%s' not numeric)", r, seq)
            continue
        if seq_int is None:
            logger.info("  row %d: skip (编号 empty)", r)
            continue
        raw_price = ws.cell(r, COL_PRICE).value
        raw_ship = ws.cell(r, COL_SHIPPING).value
        try:
            price = float(raw_price) if raw_price not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (原价 '%s' not numeric)", r, raw_price)
            continue
        try:
            shipping = float(raw_ship) if raw_ship not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (运费 '%s' not numeric)", r, raw_ship)
            continue
        variant_filter = str(ws.cell(r, COL_VARIANT_FILTER).value or "").strip()
        pending.append({
            "row": r,
            "seq": seq_int,
            "url": str(url).strip(),
            "price": price,
            "shipping": shipping,
            "variant_filter": variant_filter,
        })
    return pending


def _ensure_enriched_sheet(enriched_path, sheet_name: str):
    """Load or create the enriched workbook, then load or create the store
    sheet with the 21-col header row. Returns (wb, ws). Raises ValueError if
    a sheet exists with mismatched headers (prevents silent schema drift)."""
    from pathlib import Path
    from openpyxl import Workbook, load_workbook

    p = Path(enriched_path)
    if p.exists():
        wb = load_workbook(p)
    else:
        wb = Workbook()
        # openpyxl seeds a default 'Sheet' we don't want in the final layout.
        default_name = wb.sheetnames[0]
        if default_name != sheet_name:
            del wb[default_name]

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Validate headers on existing sheet — hard error on drift.
        for ci, expected in enumerate(EXPECTED_HEADERS, 1):
            actual = ws.cell(1, ci).value
            if actual != expected:
                raise ValueError(
                    f"Enriched sheet '{sheet_name}' header mismatch at col {ci}: "
                    f"expected {expected!r}, got {actual!r}. Fix the file or "
                    f"delete the sheet to have it recreated."
                )
    else:
        ws = wb.create_sheet(sheet_name)
        for ci, header in enumerate(EXPECTED_HEADERS, 1):
            ws.cell(1, ci).value = header

    return wb, ws


def _sheet_matches_template(ws) -> bool:
    """The sheet must have '链接' in col B row 1 to be treated as the new schema."""
    return str(ws.cell(1, COL_URL).value or "").strip() == "链接"


def _read_pending_rows(ws) -> list[dict]:
    """Return [{row, seq, url, price, shipping, variant_filter}, ...] for
    rows that have 链接 filled and 日期/状态 both empty. Non-template sheets
    return []. Empty 原价/运费 come through as None (scraper fallback)."""
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        url = ws.cell(r, COL_URL).value
        date_v = ws.cell(r, COL_DATE).value
        status_v = ws.cell(r, COL_STATUS).value
        if not url or date_v or status_v:
            continue
        if seq is None:
            logger.info("  row %d: skip (no 编号)", r)
            continue
        raw_price = ws.cell(r, COL_PRICE).value
        raw_ship = ws.cell(r, COL_SHIPPING).value
        try:
            price = float(raw_price) if raw_price not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (原价 '%s' not numeric)", r, raw_price)
            continue
        try:
            shipping = float(raw_ship) if raw_ship not in (None, "") else None
        except (TypeError, ValueError):
            logger.info("  row %d: skip (运费 '%s' not numeric)", r, raw_ship)
            continue
        try:
            seq_int = int(seq)
        except (TypeError, ValueError):
            logger.info("  row %d: skip (编号 '%s' not numeric)", r, seq)
            continue
        variant_filter = str(ws.cell(r, COL_VARIANT_FILTER).value or "").strip()
        pending.append({
            "row": r,
            "seq": seq_int,
            "url": str(url).strip(),
            "price": price,
            "shipping": shipping,
            "variant_filter": variant_filter,
        })
    return pending


def _write_result_row(
    ws, row: int, date: str, status: str,
    picture_path=None,
    web_price=None,
    shein_title=None,
    ebay_title=None,
    ebay_price=None,
    stock=None,
) -> None:
    """Write result columns F–M. Optional cols default to None (skip write).
    Always writes 日期 (F) and 状态 (G). When picture_path is a real file,
    embeds the image in H and grows the row height to fit."""
    ws.cell(row, COL_DATE).value = date
    ws.cell(row, COL_STATUS).value = status
    if picture_path and Path(picture_path).is_file():
        row_h = _add_picture_to_cell(ws, row, COL_PICTURE, Path(picture_path))
        if row_h:
            existing = ws.row_dimensions[row].height or 0
            ws.row_dimensions[row].height = max(existing, row_h)
    if web_price is not None:
        ws.cell(row, COL_WEB_PRICE).value = web_price
    if shein_title is not None:
        ws.cell(row, COL_SHEIN_TITLE).value = shein_title
    if ebay_title is not None:
        ws.cell(row, COL_EBAY_TITLE).value = ebay_title
    if ebay_price is not None:
        ws.cell(row, COL_EBAY_PRICE).value = ebay_price
    if stock is not None:
        ws.cell(row, COL_STOCK).value = stock


def setup_logging():
    DEBUG_LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = DEBUG_LOG_DIR / f"excel_{ts}.log"
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    root = logging.getLogger("run_excel")
    root.handlers.clear()
    root.setLevel(logging.DEBUG)

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    sh.setFormatter(fmt)
    root.addHandler(sh)

    logger.info("Log: %s", log_path)
    return log_path


def safe_save(wb, xlsx_path: Path) -> None:
    """Save workbook. If locked by another user, save as copy with '2' suffix."""
    try:
        wb.save(xlsx_path)
    except PermissionError:
        alt = xlsx_path.with_stem(xlsx_path.stem + "2")
        logger.warning("Cannot save to %s (locked), saving to %s", xlsx_path.name, alt.name)
        wb.save(alt)
        logger.info("Saved to alternate: %s", alt.name)


def process_excel(xlsx_path: Path) -> None:
    """Process all worksheets in `xlsx_path` using the Chinese-header template
    schema. Non-template sheets are skipped. See design spec §1 for columns."""
    logger.info("Opening: %s", xlsx_path.name)
    wb = load_workbook(xlsx_path)

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        store = ws_name.strip()
        pending = _read_pending_rows(ws)
        if not pending:
            logger.info("  Sheet '%s': no pending rows (or not template schema)", store)
            continue

        logger.info("  Sheet '%s': %d pending row(s): seq %s",
                    store, len(pending), [p["seq"] for p in pending])

        store_dir = OUTPUT_ROOT / store
        store_dir.mkdir(parents=True, exist_ok=True)

        today = datetime.now().strftime("%Y-%m-%d")
        seqs = [p["seq"] for p in pending]
        seq_min, seq_max = min(seqs), max(seqs)
        batch_xlsx = f"{store}-{seq_min}-{seq_max}-{today.replace('-', '')}.xlsx"

        urls = [p["url"] for p in pending]
        prices = [p["price"] for p in pending]
        shippings = [p["shipping"] for p in pending]
        variant_filters = [p["variant_filter"] for p in pending]

        old_cwd = Path.cwd()
        results = None
        try:
            for _retry in range(5):
                try:
                    os.chdir(store_dir)
                    break
                except PermissionError:
                    time.sleep(2)
            else:
                os.chdir(store_dir)
            logger.info("  Scraping %d URLs → %s/%s", len(urls), store, batch_xlsx)
            results = scrape_shein(
                urls,
                output=batch_xlsx,
                seq_list=seqs,
                price_list=prices,
                shipping_list=shippings,
                variant_filter_list=variant_filters,
            )
        except RateLimitError:
            logger.warning("  [限流] Rate limited during '%s'", store)
        except Exception as e:
            logger.exception("  Error processing '%s': %s", store, e)
            try:
                tb = traceback.format_exc()
                (DEBUG_LOG_DIR / "last_traceback.txt").write_text(tb, encoding="utf-8")
            except OSError:
                pass
        finally:
            os.chdir(old_cwd)

        # Fan results back into the template rows
        for p in pending:
            row = p["row"]
            seq = p["seq"]
            seq_folder = store_dir / str(seq)
            has_files = seq_folder.is_dir() and any(seq_folder.iterdir())

            rec = None
            if results:
                rec = next((r for r in results if r.get("seq_num") == seq), None)

            is_bad_data = (rec and rec.get("status") == "OK"
                           and (not rec.get("sku")
                                or "[goods_name]" in (rec.get("title") or "")))

            if has_files and not is_bad_data and rec:
                _write_result_row(
                    ws, row=row,
                    date=today, status="Done",
                    picture_path=rec.get("first_image_path") or None,
                    web_price=rec.get("web_price_display"),
                    shein_title=rec.get("original_title") or rec.get("title"),
                    ebay_title=rec.get("ebay_title"),
                    ebay_price=rec.get("ebay_price"),
                    stock=rec.get("stock_summary"),
                )
                logger.info("    seq %d → Done", seq)
            elif rec and rec.get("status") == "DELISTED":
                _write_result_row(ws, row=row, date=today, status="Delisted")
                logger.info("    seq %d → Delisted", seq)
            else:
                detail = rec.get("status", "") if rec else ""
                if is_bad_data:
                    detail = "no data loaded"
                _write_result_row(ws, row=row, date=today, status="Failed")
                logger.info("    seq %d → Failed %s", seq,
                            f"({detail})" if detail else "")

        safe_save(wb, xlsx_path)
        logger.info("  Saved progress to %s", xlsx_path.name)

    wb.close()
    logger.info("Done: %s", xlsx_path.name)


def _discover_xlsx(submitted_dir: Path) -> list[Path]:
    """Top-level .xlsx files only (so the 上架资料-已完成 subfolder isn't scanned).
    Skip Excel temp lock files (~$...)."""
    if not submitted_dir.is_dir():
        return []
    files = []
    for p in submitted_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() != ".xlsx":
            continue
        if p.name.startswith("~$"):
            continue
        files.append(p)
    return sorted(files)


def main():
    parser = argparse.ArgumentParser(
        description="Excel-based Shein scraper pipeline (澳洲站)")
    parser.add_argument("file", nargs="?", default=None,
                        help="Path to .xlsx file (default: 处理 SUBMITTED_DIR 下所有 .xlsx)")
    args = parser.parse_args()

    setup_logging()

    if args.file:
        files = [Path(args.file)]
    elif INPUT_FILENAME:
        candidate = SUBMITTED_DIR / INPUT_FILENAME
        if candidate.exists():
            files = [candidate]
        else:
            logger.error("INPUT_FILENAME 设置但找不到: %s", candidate)
            return
    else:
        files = _discover_xlsx(SUBMITTED_DIR)
        if not files:
            logger.error("SUBMITTED_DIR 下没有 .xlsx 文件: %s", SUBMITTED_DIR)
            return
        logger.info("发现 %d 个输入文件: %s", len(files), [f.name for f in files])

    for f in files:
        logger.info("=" * 60)
        try:
            process_excel(f)
        except Exception as e:
            logger.exception("Fatal error processing %s: %s", f.name, e)

    logger.info("All done.")


if __name__ == "__main__":
    main()
