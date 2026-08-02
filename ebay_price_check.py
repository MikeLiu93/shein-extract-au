"""Post-scrape helper: for each row where G=Done and N=empty, search
ebay.com.au via CDP browser (3-tab parallel), extract top-2 Best Match
listings, and write delivered price + cleaned URL to columns N-R.

Usage:
    python ebay_price_check.py                  # picks SUBMITTED_DIR/SHEIN_INPUT_FILENAME
    python ebay_price_check.py "path/to/x.xlsx" # explicit file

Columns written (see docs/superpowers/specs/2026-08-02-ebay-price-check-design.md):
    N: eBay 搜索日期    — YYYY-MM-DD
    O: eBay 同类低价    — numeric (rank-1 or rank-2, whichever is cheaper)
    P: 低价链接         — cleaned eBay itm URL
    Q: eBay 同类高价    — numeric (the pricier of the two)
    R: 高价链接         — cleaned eBay itm URL

Row selection: G exactly 'Done' AND N empty AND K (eBay 标题) non-empty.
"""

import argparse
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

# Column indices for the eBay result block.
COL_EBAY_SEARCH_DATE = 14
COL_LOW_PRICE        = 15
COL_LOW_URL          = 16
COL_HIGH_PRICE       = 17
COL_HIGH_URL         = 18

logger = logging.getLogger("ebay_price_check")
DEBUG_LOG_DIR = Path(__file__).resolve().parent / "debug_logs"


def _read_ebay_pending_rows(ws) -> list:
    """Return list of dicts for rows that need an eBay search:
    G == 'Done' AND N empty AND K (eBay 标题) non-empty. Non-template
    sheets (missing '链接' in B1) return [].
    Dict shape: {row, seq, ebay_title, shein_title}."""
    # Import here to avoid circular deps at module load time.
    from run_excel import _sheet_matches_template, COL_SEQ, COL_STATUS, COL_SHEIN_TITLE, COL_EBAY_TITLE
    if not _sheet_matches_template(ws):
        return []
    pending = []
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(r, COL_SEQ).value
        status = ws.cell(r, COL_STATUS).value
        search_date = ws.cell(r, COL_EBAY_SEARCH_DATE).value
        ebay_title = ws.cell(r, COL_EBAY_TITLE).value
        shein_title = ws.cell(r, COL_SHEIN_TITLE).value
        if str(status or "") != "Done":
            continue
        if search_date not in (None, ""):
            continue
        if not (ebay_title and str(ebay_title).strip()):
            logger.info("  row %d: skip (eBay 标题 empty)", r)
            continue
        try:
            seq_int = int(seq) if seq is not None else None
        except (TypeError, ValueError):
            seq_int = None
        pending.append({
            "row": r,
            "seq": seq_int,
            "ebay_title": str(ebay_title).strip(),
            "shein_title": str(shein_title or "").strip(),
        })
    return pending


def _write_ebay_result(
    ws, row: int, search_date: str,
    low_price=None, low_url: str = None,
    high_price=None, high_url: str = None,
    no_match: bool = False,
) -> None:
    """Write columns N-R for one row. Always writes N (search_date).
    On no_match=True: O = 'no match', P/Q/R empty. Otherwise: fill the
    slots provided; unspecified stays None."""
    ws.cell(row, COL_EBAY_SEARCH_DATE).value = search_date
    if no_match:
        ws.cell(row, COL_LOW_PRICE).value = "no match"
        return
    if low_price is not None:
        ws.cell(row, COL_LOW_PRICE).value = low_price
    if low_url:
        ws.cell(row, COL_LOW_URL).value = low_url
    if high_price is not None:
        ws.cell(row, COL_HIGH_PRICE).value = high_price
    if high_url:
        ws.cell(row, COL_HIGH_URL).value = high_url
