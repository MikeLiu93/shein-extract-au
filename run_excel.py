"""
Excel-based pipeline (澳洲站): read pending URLs from .xlsx worksheets (one per
store), scrape them, write results to OUTPUT_ROOT/{store}/, and update Date +
Status columns in the source Excel.

Usage:
    python run_excel.py                          # 扫描 SUBMITTED_DIR 下所有 .xlsx
    python run_excel.py "path/to/file.xlsx"      # 指定文件

Columns (strict):
    A: Seq       — sequence number (= output folder name). READ ONLY.
    B: Website   — Shein product URL. READ ONLY.
    C: Price     — manual product price (USD). READ ONLY. Overrides web sale_price.
    D: Date      — filled after run: YYYY-MM-DD. WRITE.
    E: Status    — filled after run: Done / Failed / Delisted. WRITE.
    H: Web price — web-scraped sale price, for manual comparison with C. WRITE.

Only rows where Website AND Price are both filled (and Date+Status both empty)
are processed. Rows with an empty Price are skipped — Price doubles as the
"this row is ready" flag.
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

from shein_scraper import scrape_shein, RateLimitError
from config import SUBMITTED_DIR, OUTPUT_ROOT_2ND as OUTPUT_ROOT, INPUT_FILENAME

logger = logging.getLogger("run_excel")
DEBUG_LOG_DIR = Path(__file__).resolve().parent / "debug_logs"


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


def process_excel(xlsx_path: Path, gate_price: bool = True) -> None:
    """Process all worksheets in an Excel file.

    gate_price=True (default): only rows with both Website AND Price filled
    are processed; empty-Price rows are skipped. gate_price=False: empty-Price
    rows are also processed, using the web-scraped price (no manual override).
    """
    logger.info("Opening: %s  (gate_price=%s)", xlsx_path.name, gate_price)
    wb = load_workbook(xlsx_path)

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        store = ws_name.strip()
        logger.info("Sheet: %s", store)

        # Normalize headers: col C = Price, col D = Date
        if str(ws.cell(1, 3).value or "").strip() == "":
            ws.cell(1, 3).value = "Price"
        header_d = str(ws.cell(1, 4).value or "").strip()
        if header_d in ("Execute Date", ""):
            ws.cell(1, 4).value = "Date"
        if str(ws.cell(1, 8).value or "").strip() == "":
            ws.cell(1, 8).value = "Web price"

        # Collect pending rows. gate_price=True (default): require Price filled
        # (Price doubles as the "ready to list" flag). gate_price=False: accept
        # empty Price; price_f=None for those rows so the web price is used.
        pending = []
        for row in range(2, ws.max_row + 1):
            seq = ws.cell(row, 1).value
            url = ws.cell(row, 2).value
            price_val = ws.cell(row, 3).value
            date_val = ws.cell(row, 4).value
            status_val = ws.cell(row, 5).value
            if not (url and seq is not None and not date_val and not status_val):
                continue
            if price_val in (None, ""):
                if gate_price:
                    continue
                price_f = None  # use web-scraped price
            else:
                try:
                    price_f = float(price_val)
                except (TypeError, ValueError):
                    logger.info("    seq %s → skip (Price '%s' not numeric)",
                                seq, price_val)
                    continue
            pending.append((row, int(seq), str(url).strip(), price_f))

        if not pending:
            logger.info("  No pending rows in '%s'", store)
            continue

        logger.info("  %d pending URL(s): seq %s",
                     len(pending), [p[1] for p in pending])

        # Prepare output folder
        store_dir = OUTPUT_ROOT / store
        store_dir.mkdir(parents=True, exist_ok=True)

        # Run scraper
        urls = [p[2] for p in pending]
        seqs = [p[1] for p in pending]
        prices = [p[3] for p in pending]
        today = datetime.now().strftime("%Y-%m-%d")
        seq_min, seq_max = min(seqs), max(seqs)
        xlsx_name = f"{store}-{seq_min}-{seq_max}-{today.replace('-', '')}.xlsx"

        old_cwd = Path.cwd()
        results = None
        try:
            # Google Drive sync may briefly lock new folders
            for _retry in range(5):
                try:
                    os.chdir(store_dir)
                    break
                except PermissionError:
                    time.sleep(2)
            else:
                os.chdir(store_dir)  # final attempt, let it raise
            logger.info("  Scraping %d URLs → %s/%s", len(urls), store, xlsx_name)
            results = scrape_shein(urls, output=xlsx_name, seq_list=seqs,
                                   price_list=prices)
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

        # Update Date + Status based on results
        for row_idx, seq, url, _price in pending:
            seq_folder = store_dir / str(seq)
            has_files = seq_folder.is_dir() and any(seq_folder.iterdir())

            rec = None
            if results:
                rec = next((r for r in results if r.get("seq_num") == seq), None)

            # Web price (col H): the web-scraped sale price, for manual
            # comparison against the manual Price in col C.
            if rec and rec.get("web_price") is not None:
                ws.cell(row_idx, 8).value = rec["web_price"]

            is_bad_data = (rec and rec.get("status") == "OK"
                           and (not rec.get("sku")
                                or "[goods_name]" in (rec.get("title") or "")))

            if has_files and not is_bad_data:
                ws.cell(row_idx, 4).value = today
                ws.cell(row_idx, 5).value = "Done"
                logger.info("    seq %d → Done", seq)
            elif rec and rec.get("status") == "DELISTED":
                ws.cell(row_idx, 4).value = today
                ws.cell(row_idx, 5).value = "Delisted"
                logger.info("    seq %d → Delisted", seq)
            else:
                detail = rec.get("status", "") if rec else ""
                if is_bad_data:
                    detail = "no data loaded"
                ws.cell(row_idx, 4).value = today
                ws.cell(row_idx, 5).value = "Failed"
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
    parser.add_argument("--no-price-gate", action="store_true",
                        help="处理 C 列 Price 为空的行(用网页价,不做覆盖);"
                             "默认要求 Price 填写才跑")
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
            process_excel(f, gate_price=not args.no_price_gate)
        except Exception as e:
            logger.exception("Fatal error processing %s: %s", f.name, e)

    logger.info("All done.")


if __name__ == "__main__":
    main()
