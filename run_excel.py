"""
Excel-based pipeline (澳洲站): read pending URLs from a single input .xlsx and
write scrape results back to the SAME workbook (F-M columns) in place. Before
any modification we back up the workbook to <BACKUP_DIR>/<name>-<ts>.xlsx so a
bad run can never destroy days of accumulated data.

Trigger (matches 美国站):
    A row is pending iff B (链接) is non-empty AND F (日期) + G (状态) are both
    empty. Done / Failed / Delisted rows are inert until the operator clears
    F+G manually — Failed rows do NOT auto-retry on the next run.

Usage:
    python run_excel.py                          # SUBMITTED_DIR/SHEIN_INPUT_FILENAME
    python run_excel.py "path/to/input.xlsx"     # explicit path

Schema (13 cols, matches 美国站):
    A: 编号     B: 链接     C: 原价     D: 运费     E: 变体
    F: 日期     G: 状态     H: 图片     I: 希音价格  J: 希音标题
    K: eBay标题  L: eBay价格  M: 库存

C (原价) and D (运费), if filled, override scraped values.
Only sheets with B1='链接' are processed.
"""

import argparse
import logging
import os
import shutil
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from shein_scraper import (scrape_shein, RateLimitError, _add_picture_to_cell,
                           save_workbook_atomic)
from config import (SUBMITTED_DIR, OUTPUT_ROOT_2ND as OUTPUT_ROOT,
                    INPUT_FILENAME, BACKUP_DIR)

logger = logging.getLogger("run_excel")
DEBUG_LOG_DIR = Path(__file__).resolve().parent / "debug_logs"

# ── Schema: A–M (13 cols) ─────────────────────────────────────────────────────
COL_SEQ, COL_URL, COL_PRICE, COL_SHIPPING, COL_VARIANT_FILTER = 1, 2, 3, 4, 5
COL_DATE, COL_STATUS, COL_PICTURE = 6, 7, 8
COL_WEB_PRICE, COL_SHEIN_TITLE = 9, 10
COL_EBAY_TITLE, COL_EBAY_PRICE, COL_STOCK = 11, 12, 13

# Chinese headers matching 美国站.
EXPECTED_HEADERS = ["编号", "链接", "原价", "运费", "变体",
                    "日期", "状态", "图片", "希音价格", "希音标题",
                    "eBay标题", "eBay价格", "库存"]

# 备份保留最近多少份；旧的自动删。SHEIN_BACKUP_KEEP env var 可覆盖。
BACKUP_KEEP = int(os.environ.get("SHEIN_BACKUP_KEEP", "20"))


def _sheet_matches_template(ws) -> bool:
    """The sheet must have '链接' in col B row 1 to be treated as our schema."""
    return str(ws.cell(1, COL_URL).value or "").strip() == "链接"


def _read_pending_rows(ws) -> list[dict]:
    """A row is pending iff B (链接) is non-empty AND F (日期) + G (状态) are
    both empty. Returns list of dicts with row/seq/url/price/shipping/variant_filter.

    Rows with non-numeric 编号 / 原价 / 运费 are logged and skipped rather than
    aborting the whole run.
    """
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        url = ws.cell(r, COL_URL).value
        if url in (None, ""):
            continue
        date = ws.cell(r, COL_DATE).value
        status = ws.cell(r, COL_STATUS).value
        if (date not in (None, "")) or (status not in (None, "")):
            continue

        seq = ws.cell(r, COL_SEQ).value
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
            "row": r, "seq": seq_int, "url": str(url).strip(),
            "price": price, "shipping": shipping,
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


# ── Backup + legacy migration ─────────────────────────────────────────────────

def _backup_workbook(input_path: Path, backup_dir: Path) -> Path:
    """Copy input_path into backup_dir with a timestamp suffix; return the copy path.
    Runs before any in-place modification so a botched save can be reversed."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dst = backup_dir / f"{input_path.stem}-{ts}{input_path.suffix}"
    shutil.copy2(input_path, dst)
    return dst


def _prune_backups(backup_dir: Path, stem: str, keep: int) -> int:
    """Keep the newest `keep` backups matching stem-*.xlsx; delete the rest.
    Returns number of files deleted. Legacy-* backups (from the one-shot
    migration) are excluded — those stay forever until the operator clears
    them by hand."""
    if not backup_dir.is_dir() or keep is None or keep <= 0:
        return 0
    all_bk = sorted(
        (p for p in backup_dir.glob(f"{stem}-*.xlsx")
         if "-legacy-" not in p.name),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    removed = 0
    for old in all_bk[keep:]:
        try:
            old.unlink()
            removed += 1
        except OSError:
            pass
    return removed


def _archive_legacy_enriched(old_path: Path, backup_dir: Path) -> Path:
    """Move the old 富表 (enriched output) file into backup_dir with a
    -legacy-<date>.xlsx suffix. Runs once, on the first launch after
    upgrading from the master/enriched-split architecture. Suffix collisions
    (multiple upgrades same day) get -1, -2, ... appended."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d")
    base = backup_dir / f"{old_path.stem}-legacy-{ts}{old_path.suffix}"
    dst = base
    counter = 1
    while dst.exists():
        counter += 1
        dst = backup_dir / f"{old_path.stem}-legacy-{ts}-{counter}{old_path.suffix}"
    shutil.move(str(old_path), str(dst))
    logger.warning(
        "  [升级迁移] 检测到旧富表 %s；已挪到 %s。新架构下输入表本身就是输出表。",
        old_path.name, dst,
    )
    return dst


# ── Logging / save ────────────────────────────────────────────────────────────

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
    """原子保存。被别人锁住（Excel 开着）时退回 '2' 后缀的副本。"""
    try:
        save_workbook_atomic(wb, xlsx_path)
    except PermissionError:
        alt = xlsx_path.with_stem(xlsx_path.stem + "2")
        logger.warning("Cannot save to %s (locked), saving to %s",
                       xlsx_path.name, alt.name)
        save_workbook_atomic(wb, alt)
        logger.info("Saved to alternate: %s", alt.name)


# ── Main pipeline ─────────────────────────────────────────────────────────────

def process_excel(input_path: Path) -> None:
    """Read pending rows from the input xlsx, scrape, write F–M back in place.
    Backup runs before the first write so a bad save can be reversed."""
    logger.info("Input:   %s", input_path.name)

    if not input_path.exists():
        logger.error("Input file not found: %s", input_path)
        return

    backup_path = _backup_workbook(input_path, BACKUP_DIR)
    logger.info("Backup:  %s", backup_path)
    removed = _prune_backups(BACKUP_DIR, input_path.stem, keep=BACKUP_KEEP)
    if removed:
        logger.info("  [备份] Pruned %d old backup(s), keeping newest %d",
                    removed, BACKUP_KEEP)

    wb = load_workbook(input_path)
    modified = False

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        store = ws_name.strip()
        pending = _read_pending_rows(ws)
        if not pending:
            logger.info("  Sheet '%s': no pending rows (or not scrape schema)", store)
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

        # Write results back to input xlsx in-place, one row per pending.
        for p in pending:
            seq = p["seq"]
            row = p["row"]
            seq_folder = store_dir / str(seq)
            has_files = seq_folder.is_dir() and any(seq_folder.iterdir())

            rec = None
            if results:
                rec = next((r for r in results if r.get("seq_num") == seq), None)

            is_bad_data = (rec and rec.get("status") == "OK"
                           and (not rec.get("sku")
                                or "[goods_name]" in (rec.get("title") or "")))

            if has_files and not is_bad_data and rec:
                picture = rec.get("first_image_path") or None
                _write_result_row(
                    ws, row=row,
                    date=today, status="Done",
                    picture_path=picture,
                    web_price=rec.get("web_price_display"),
                    shein_title=rec.get("original_title") or rec.get("title"),
                    ebay_title=rec.get("ebay_title"),
                    ebay_price=rec.get("ebay_price"),
                    stock=rec.get("stock_summary"),
                )
                logger.info("    seq %d row %d → Done", seq, row)
            elif rec and rec.get("status") == "DELISTED":
                _write_result_row(ws, row=row, date=today, status="Delisted")
                logger.info("    seq %d row %d → Delisted "
                            "(clear F+G to retry)", seq, row)
            else:
                detail = rec.get("status", "") if rec else ""
                if is_bad_data:
                    detail = "no data loaded"
                _write_result_row(ws, row=row, date=today, status="Failed")
                logger.info("    seq %d row %d → Failed %s "
                            "(clear F+G to retry)", seq, row,
                            f"({detail})" if detail else "")
            modified = True

    if modified:
        safe_save(wb, input_path)
        logger.info("Saved results to %s", input_path.name)
    else:
        logger.info("No pending rows anywhere; input untouched.")


def _one_shot_legacy_migration(input_path: Path) -> None:
    """If an old 富表 (from the pre-single-workbook architecture) is still
    sitting in SUBMITTED_DIR, move it to the backup folder so employees don't
    stare at two parallel workbooks. Detects the old file via the SHEIN_OUTPUT_FILENAME
    env var (still present in employees' legacy config.env from 0.3.6 and earlier)."""
    old_name = os.environ.get("SHEIN_OUTPUT_FILENAME", "").strip()
    if not old_name:
        return
    old_p = SUBMITTED_DIR / old_name
    if not old_p.exists():
        return
    try:
        # Don't move the input xlsx if the user reused the same filename.
        if old_p.resolve() == input_path.resolve():
            return
    except OSError:
        return
    try:
        _archive_legacy_enriched(old_p, BACKUP_DIR)
    except OSError as e:
        logger.warning("  [升级迁移] 无法搬走旧富表 %s: %s", old_p, e)


def main():
    parser = argparse.ArgumentParser(
        description="Single-workbook Shein scrape pipeline (澳洲站)")
    parser.add_argument("input_file", nargs="?", default=None,
                        help="Input .xlsx path. Default: "
                             "SUBMITTED_DIR/SHEIN_INPUT_FILENAME.")
    args = parser.parse_args()

    setup_logging()

    if args.input_file:
        input_path = Path(args.input_file)
    elif INPUT_FILENAME:
        input_path = SUBMITTED_DIR / INPUT_FILENAME
    else:
        logger.error("No input file given and SHEIN_INPUT_FILENAME not set in config.")
        return

    _one_shot_legacy_migration(input_path)

    logger.info("=" * 60)
    try:
        process_excel(input_path)
    except Exception as e:
        logger.exception("Fatal error: %s", e)
    logger.info("All done.")


if __name__ == "__main__":
    main()
