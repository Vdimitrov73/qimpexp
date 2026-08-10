"""
qif_importer.py  —  Phase 2: Import Quicken QIF back into ACB workbook.

Supported QIF actions  (controlled by mode parameter):
  full             Buy + Sell + RtrnCapX (ROC)
  buys-sells       Buy + Sell only
  tax-adjustments  RtrnCapX (ROC) only

Date conversion rules for Canadian securities (TSX, T+1 settlement):
  - Buy / Sell: QIF has trade date → ACB wants settlement date (next_trading_day)
  - RtrnCapX:   QIF has record date = ACB record date → no conversion

Quicken native QIF date format M/D'YY and standard MM/DD/YYYY are both handled.

Formatting: every written cell matches the existing spreadsheet style exactly —
  same number formats, font, alignment, and thin black borders as the rows above.
  Styles are copied from the nearest existing data row (not hard-coded), so the
  importer works correctly across different workbook versions.

Insertion order: rows are inserted in date order using ws.insert_rows(), exactly
  as update_acb.py does it. This means importing 2024 transactions into a sheet
  that already has 2025 rows works correctly.
"""

import re
import shutil
import tempfile
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from ca_calendar import next_trading_day

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False


# ---------------------------------------------------------------------------
# Formatting constants — must match acb_worksheet_template
# ---------------------------------------------------------------------------

DATE_FMT    = "yyyy\\-mmm\\-dd"         # matches existing workbook exactly
PRICE_FMT   = '"$"#,##0.#######;\\-"$"#,##0.#######'
SHARES_FMT  = "General"                 # col D: blank on ROC rows, number on Buy/Sell
COMM_FMT    = '"$"#,##0.00;\\-"$"#,##0.00'
CAPGAIN_FMT = '\\$#,##0.00;"-$"#,##0.00;\\$0.00'
BALANCE_FMT = "#,##0.####"
ACBCHG_FMT  = '\\$#,##0.00;"-$"#,##0.00;\\$0.00'
ACB_FMT     = '\\$#,##0.00;"-$"#,##0.00;\\$0.00'
ACBSHR_FMT  = '\\$#,##0.######;"-$"#,##0.######'

# Column index → number format (1-based)
_COL_FMT = {
    1: DATE_FMT,
    2: "General",
    3: PRICE_FMT,
    4: SHARES_FMT,
    5: COMM_FMT,
    6: CAPGAIN_FMT,
    7: BALANCE_FMT,
    8: ACBCHG_FMT,
    9: ACB_FMT,
    10: ACBSHR_FMT,
}

_thin_black = Side(style="thin", color="000000")
_BORDER     = Border(
    top=_thin_black, bottom=_thin_black,
    left=_thin_black, right=_thin_black,
)

# Alignment per column (1-based)
_COL_ALIGN = {
    1: Alignment(horizontal="left"),
    2: Alignment(horizontal="general"),
    3: Alignment(horizontal="right"),
    4: Alignment(horizontal="right"),
    5: Alignment(horizontal="right"),
    6: Alignment(horizontal="right", vertical="center"),
    7: Alignment(horizontal="right", vertical="center"),
    8: Alignment(horizontal="right", vertical="center"),
    9: Alignment(horizontal="right", vertical="center"),
    10: Alignment(horizontal="right", vertical="center"),
}


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def _get_ref_fonts(ws, last_data_row: int) -> dict:
    """
    Read font objects from the nearest existing data row so we copy the exact
    workbook font rather than hard-coding a name/size.
    Returns {col_index: Font} for columns 1-10.
    """
    fonts = {}
    # Walk backwards from last_data_row to find a row with actual data in col B
    for r in range(last_data_row, 1, -1):
        if ws.cell(r, 2).value:
            for col in range(1, 11):
                fonts[col] = copy(ws.cell(r, col).font)
            return fonts
    # Fallback: use header row fonts
    for col in range(1, 11):
        fonts[col] = copy(ws.cell(1, col).font)
    return fonts


def _apply_cell(cell, number_format: str, font, alignment, border):
    """Apply all style properties to a single cell."""
    cell.number_format = number_format
    cell.font          = font
    cell.alignment     = alignment
    cell.border        = border


def _write_formulas(ws, r: int, fonts: dict):
    """
    Write formula cells F(6)–J(10) for row r using the same formula logic as
    update_acb.py / acb_worksheet_template, and apply correct formatting.
    """
    prev   = r - 1
    G_prev = f"G{prev}" if prev >= 2 else "0"
    I_prev = f"I{prev}" if prev >= 2 else "0"
    J_prev = f"J{prev}" if prev >= 2 else "0"

    formulas = {
        # F – Capital Gains
        6: (f'=IF(B{r}="Sell",'
            f'ROUND((C{r}-{J_prev})*D{r}-IF(ISNUMBER(E{r}),E{r},0),2),0)',
            CAPGAIN_FMT),
        # G – Share Balance
        7: (f'=ROUND(IF(B{r}="Buy",{G_prev}+D{r},'
            f'IF(B{r}="Sell",{G_prev}-D{r},{G_prev})),4)',
            BALANCE_FMT),
        # H – ACB Change
        8: (f'=IF(B{r}="Buy",'
            f'ROUND(C{r}*D{r}+IF(ISNUMBER(E{r}),E{r},0),2),'
            f'IF(B{r}="Sell",'
            f'ROUND(-({J_prev}*D{r}),2),'
            f'IF(B{r}="ROC",'
            f'ROUND(-C{r}*{G_prev},2),0)))',
            ACBCHG_FMT),
        # I – Cumulative ACB
        9:  (f'=ROUND({I_prev}+H{r},2)', ACB_FMT),
        # J – ACB / Share
        10: (f'=IF(G{r}<>0,ROUND(I{r}/G{r},7),0)', ACBSHR_FMT),
    }

    for col, (formula, fmt) in formulas.items():
        cell = ws.cell(r, col, formula)
        _apply_cell(cell, fmt, fonts.get(col, Font()), _COL_ALIGN[col], _BORDER)


def _rewrite_formulas_from(ws, start_row: int, last_data_row: int, fonts: dict):
    """Rewrite all formula rows from start_row downward after an insertion."""
    for r in range(start_row, last_data_row + 1):
        _write_formulas(ws, r, fonts)


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5,  "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _parse_cell_date(val) -> date | None:
    """Parse a worksheet cell value (datetime, date, or string) to a date."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    # YYYY-MMM-DD  (workbook format)
    m = re.match(r"(\d{4})-([A-Za-z]{3})-(\d{2})", s)
    if m:
        return date(int(m.group(1)), MONTH_MAP[m.group(2).lower()], int(m.group(3)))
    # YYYY-MM-DD
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _find_last_data_row(ws) -> int:
    """Return 1-based index of the last row that has a value in col B."""
    last = 1
    for r in range(2, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if val and str(val).strip() in ("Buy", "Sell", "ROC"):
            last = r
    return last


def _find_insertion_row(ws, target_date: date, last_data_row: int) -> int:
    """
    Return the row number to INSERT BEFORE, so the new row ends up after
    all existing rows with the same or earlier date.  Mirrors update_acb.py.
    """
    insert_after = 1
    for r in range(2, last_data_row + 1):
        d = _parse_cell_date(ws.cell(r, 1).value)
        if d is not None and d <= target_date:
            insert_after = r
    return insert_after + 1


def _get_existing_rows(ws, last_data_row: int) -> set:
    """
    Return set of (date, type_str, round(price,7), round(shares,4)) for duplicate detection.
    """
    existing = set()
    for r in range(2, last_data_row + 1):
        txn_type = ws.cell(r, 2).value
        if not txn_type:
            continue
        d = _parse_cell_date(ws.cell(r, 1).value)
        price_val  = ws.cell(r, 3).value
        shares_val = ws.cell(r, 4).value
        if d and txn_type and price_val is not None:
            try:
                typ = str(txn_type).strip()
                if typ in ("Buy", "Sell"):
                    pk = round(float(price_val), 2)   # fuzzy: catches Quicken truncation
                    sk = round(float(shares_val), 4) if shares_val is not None else None
                else:                                  # ROC
                    pk = round(float(price_val), 5)   # precise enough to distinguish same-date ROCs
                    sk = None                          # ROC rows have no shares in col D
                existing.add((d, typ, pk, sk))
            except (ValueError, TypeError):
                pass
    return existing

def _format_date_str(d: date) -> str:
    """Format a date as YYYY-MMM-DD (workbook canonical format)."""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{d.year}-{months[d.month - 1]}-{d.day:02d}"


def _get_share_balance_before_row(ws, insert_row: int):
    """
    Compute share balance by accumulating Buy/Sell shares from col D
    up to (but not including) insert_row.  Reads input columns B and D
    only — never formula columns — so it works with both normal and
    data_only workbook loads.
    """
    balance   = Decimal("0")
    found_any = False
    for r in range(2, insert_row):
        typ = ws.cell(r, 2).value
        if not typ:
            continue
        typ = str(typ).strip()
        shares_val = ws.cell(r, 4).value
        if typ in ("Buy", "Sell") and shares_val is not None:
            try:
                delta = Decimal(str(shares_val))
                balance += delta if typ == "Buy" else -delta
                found_any = True
            except Exception:
                pass
        # ROC rows intentionally skipped — they don't affect share count
    return balance if found_any else None

def _derive_roc_price(amount: Decimal, share_balance: Decimal):
    """
    Derive per-unit ROC price from total distribution amount and share balance.
    Tries 4-8 decimal places and returns the first precision that exactly
    reconstructs the 2dp amount.  Falls back to 7dp if none match exactly.
    """
    if share_balance == 0:
        return None
    raw = amount / share_balance
    for dp in range(4, 9):
        q = Decimal("0." + "0" * dp)
        candidate = raw.quantize(q, rounding=ROUND_HALF_UP)
        reconstructed = (candidate * share_balance).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        if reconstructed == amount:
            return candidate
    # No exact match — return 7dp best approximation
    return raw.quantize(Decimal("0.0000001"), rounding=ROUND_HALF_UP)


# ---------------------------------------------------------------------------
# QIF parsing
# ---------------------------------------------------------------------------

_RE_QUICKEN_DATE = re.compile(r"^(\d{1,2})/\s*(\d{1,2})'(\d{2})$")
_RE_STD_DATE     = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")


def _parse_qif_date(s: str) -> date | None:
    s = s.strip()
    m = _RE_QUICKEN_DATE.match(s)
    if m:
        month, day, year2 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + year2 if year2 < 50 else 1900 + year2
        return date(year, month, day)
    m = _RE_STD_DATE.match(s)
    if m:
        return date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _parse_qif_decimal(s: str) -> Decimal | None:
    if not s:
        return None
    try:
        s = s.strip().replace(",", "").replace("(", "-").replace(")", "")
        return Decimal(s)
    except InvalidOperation:
        return None


def parse_qif_file(qif_path: Path) -> tuple:
    """
    Parse a QIF investment file.
    Returns (account_name, transactions, skipped_actions, error_str | None).
    """
    account_name    = None
    transactions    = []
    skipped_actions = {}
    current         = {}

    # Actions the importer can write to the ACB workbook
    _WRITABLE = {"Buy", "Sell", "RtrnCapX"}

    def _flush():
        nonlocal current
        if not current:
            return
        action = current.get("action")
        if action not in _WRITABLE:
            skipped_actions[action] = skipped_actions.get(action, 0) + 1
            current = {}
            return
        if "trade_date" not in current:
            current = {}
            return
        transactions.append(dict(current))
        current = {}

    try:
        with open(qif_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
    except OSError as exc:
        return None, [], {}, str(exc)

    for raw_line in lines:
        line = raw_line.rstrip("\r\n")
        if not line:
            continue
        tag  = line[0]
        body = line[1:].strip()

        if tag == "!":
            continue
        if tag == "L" and not current:
            m = re.match(r"^\[(.+)\]$", body)
            if m:
                account_name = m.group(1)
            continue
        if tag == "^":
            _flush()
            continue
        if tag == "D":
            d = _parse_qif_date(body)
            if d:
                current["trade_date"] = d
        elif tag == "N":
            current["action"] = body
        elif tag == "Y":
            current["security_name"] = body
        elif tag == "I":
            current["price"] = _parse_qif_decimal(body)
        elif tag == "Q":
            current["shares"] = _parse_qif_decimal(body)
        elif tag in ("T", "U"):
            if "amount" not in current:
                current["amount"] = _parse_qif_decimal(body)
        elif tag == "M":
            current["memo"] = body

    _flush()
    return account_name, transactions, skipped_actions, None


# ---------------------------------------------------------------------------
# Security map helpers
# ---------------------------------------------------------------------------

def _find_ticker_for_security(security_name: str, security_map: dict) -> str | None:
    target = security_name.strip().upper()
    for ticker, name in security_map.items():
        if name.strip().upper() == target:
            return ticker
    return None


# ---------------------------------------------------------------------------
# Main import function
# ---------------------------------------------------------------------------

def import_qif_to_acb(
    qif_path: Path,
    acb_path: Path,
    security_map: dict,
    mode: str = "full",            # "full" | "buys-sells" | "tax-adjustments"
    dry_run: bool = False,
    verbose: bool = False,
) -> dict:
    """
    Read *qif_path* and insert matching transactions into *acb_path* in date order.

    mode:
      "full"             — Buy + Sell + ROC
      "buys-sells"       — Buy + Sell only
      "tax-adjustments"  — ROC (RtrnCapX) only

    Rows are inserted at the correct chronological position using ws.insert_rows(),
    so importing older transactions into a sheet with newer data works correctly.
    Formatting (number format, font, alignment, borders) matches the existing rows.
    """
    warnings_out = []

    # --- Determine which actions to import based on mode ---
    if mode == "full":
        allowed_actions = {"Buy", "Sell", "RtrnCapX"}
    elif mode == "buys-sells":
        allowed_actions = {"Buy", "Sell"}
    elif mode == "tax-adjustments":
        allowed_actions = {"RtrnCapX"}
    else:
        allowed_actions = {"Buy", "Sell", "RtrnCapX"}

    # --- Parse QIF ---
    account_name, transactions, skipped_actions, err = parse_qif_file(qif_path)
    if err:
        return {"fatal": f"Cannot read QIF: {err}", "warnings": [], "imported": 0,
                "skipped": 0, "rows_by_sheet": {}}

    if verbose:
        print(f"[QIF-IMPORT] Account:      {account_name or '(not found in file)'}")
        print(f"[QIF-IMPORT] Transactions: {len(transactions)}")
        print(f"[QIF-IMPORT] Mode:         {mode}")
        if skipped_actions:
            print(f"[QIF-IMPORT] Skipped actions: {skipped_actions}")

    if not transactions:
        return {"fatal": "No importable transactions found in QIF file.",
                "warnings": [], "imported": 0, "skipped": 0, "rows_by_sheet": {}}

    # --- Map each transaction to a ticker and build row data ---
    rows_by_sheet: dict = {}
    imported = 0
    skipped  = 0

    late_skipped = 0  # rows that pass initial mapping but fail during write

    for txn in transactions:
        action = txn["action"]

        # Filter by mode
        if action not in allowed_actions:
            if verbose:
                print(f"[MODE-SKIP] {action} not in mode={mode!r}")
            skipped += 1
            continue

        sec_name = txn.get("security_name", "")
        trade_date = txn["trade_date"]

        # ROC has no security_name in some exports — but RtrnCapX always does
        if not sec_name and action == "RtrnCapX":
            w = f"[SKIP] {trade_date} RtrnCapX: no security name in QIF — skipped"
            warnings_out.append(w)
            if verbose:
                print(w)
            skipped += 1
            continue

        ticker = _find_ticker_for_security(sec_name, security_map)
        if ticker is None:
            w = (f"[SKIP] {trade_date} {action} {sec_name!r}: "
                 f"no ticker mapping — skipped")
            warnings_out.append(w)
            if verbose:
                print(w)
            skipped += 1
            continue

        price  = txn.get("price")
        shares = txn.get("shares")
        amount = txn.get("amount")

        # --- Date conversion ---
        if action in ("Buy", "Sell"):
            acb_date = next_trading_day(trade_date)
            if verbose:
                print(f"[DATE] {ticker} {trade_date} ({action}) -> settlement {acb_date}")
        else:
            acb_date = trade_date

        # --- Derive missing price or shares from amount for Buy/Sell ---
        if action in ("Buy", "Sell"):
            if price is not None and shares is None and amount is not None:
                try:
                    shares = (amount / price).quantize(Decimal("0.0001"))
                except Exception:
                    pass
            if shares is not None and price is None and amount is not None:
                try:
                    price = (amount / shares).quantize(Decimal("0.000001"))
                except Exception:
                    pass
            if price is None or shares is None:
                w = (f"[SKIP] {ticker} {acb_date} {action}: "
                     f"cannot determine price or shares — skipped")
                warnings_out.append(w)
                if verbose:
                    print(w)
                skipped += 1
                continue

        # --- ROC: per-unit price from I line, or derived later from share balance ---
        if action == "RtrnCapX":
            acb_type  = "ROC"
            roc_price = price  # per-unit from QIF I line — may be None

            if roc_price is None and amount is None:
                w = (f"[SKIP] {ticker} {acb_date} ROC: no price or amount in QIF "
                     f"— skipped; enter col C manually")
                warnings_out.append(w)
                if verbose:
                    print(w)
                skipped += 1
                continue
            # roc_price may still be None here — will be derived from
            # workbook share balance when writing the row

            row_data = {
                "ticker":     ticker,
                "acb_date":   acb_date,
                "trade_date": trade_date,
                "acb_type":   acb_type,
                "price":      roc_price,   # None = derive from workbook
                "shares":     None,
                "commission": None,
                "raw_amount": amount,
            }

        else:
            acb_type = action
            row_data = {
                "ticker":     ticker,
                "acb_date":   acb_date,
                "trade_date": trade_date,
                "acb_type":   acb_type,
                "price":      price,
                "shares":     shares,
                "commission": None,
                "raw_amount": amount,
            }

        rows_by_sheet.setdefault(ticker, []).append(row_data)
        imported += 1

    if not rows_by_sheet:
        return {"fatal": "No transactions matched a known ticker.",
                "warnings": warnings_out, "imported": 0,
                "skipped": skipped, "rows_by_sheet": {}}

    # --- Dry run ---
    if dry_run:
        # Open workbook read-only so we can derive ROC prices for display
        try:
            wb_ro = openpyxl.load_workbook(str(acb_path), data_only=True)
        except Exception:
            wb_ro = None

        print("\n[DRY RUN] Rows that would be inserted:")
        for ticker, rows in sorted(rows_by_sheet.items()):
            ws_ro = None
            if wb_ro:
                for sn in wb_ro.sheetnames:
                    if sn.strip().upper() == ticker:
                        ws_ro = wb_ro[sn]
                        break
            print(f"  Sheet: {ticker}")
            for r in sorted(rows, key=lambda x: x["acb_date"]):
                display_price = r["price"]
                if display_price is None and ws_ro and r.get("raw_amount") is not None:
                    last = _find_last_data_row(ws_ro)
                    ins  = _find_insertion_row(ws_ro, r["acb_date"], last)
                    bal  = _get_share_balance_before_row(ws_ro, ins)
                    if bal:
                        display_price = _derive_roc_price(r["raw_amount"], bal)
                print(f"    {r['acb_date']}  {r['acb_type']:4s}  "
                      f"price={display_price}  shares={r['shares']}  "
                      f"amount={r['raw_amount']}")
        if wb_ro:
            wb_ro.close()

        return {
            "fatal": None, "warnings": warnings_out,
            "imported": imported, "skipped": skipped + late_skipped,
            "rows_by_sheet": rows_by_sheet,
            "sheets_written": {t: len(r) for t, r in rows_by_sheet.items()},
        }

    # --- Write to workbook ---
    if not _HAS_OPENPYXL:
        return {"fatal": "openpyxl required to write workbook.",
                "warnings": warnings_out, "imported": 0,
                "skipped": skipped, "rows_by_sheet": {}}

    # Backup before any modification
    acb_path = Path(acb_path)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{acb_path.stem}_backup_{timestamp}{acb_path.suffix}"
    backup_path = acb_path.with_name(backup_name)
    try:
        shutil.copy2(str(acb_path), str(backup_path))
        if verbose:
            print(f"[BACKUP] {backup_path}")
    except OSError:
        backup_path = Path(tempfile.gettempdir()) / backup_name
        try:
            shutil.copy2(str(acb_path), str(backup_path))
            warnings_out.append(f"[WARN] Backup written to temp: {backup_path}")
        except OSError as exc2:
            return {"fatal": f"Cannot create backup: {exc2}",
                    "warnings": warnings_out, "imported": 0,
                    "skipped": skipped, "rows_by_sheet": {}}

    try:
        wb = openpyxl.load_workbook(str(acb_path))
    except Exception as exc:
        return {"fatal": f"Cannot open workbook: {exc}",
                "warnings": warnings_out, "imported": 0,
                "skipped": skipped, "rows_by_sheet": {}}

    sheets_written = {}
    rollback       = False
    rollback_reason = ""

    for ticker, rows in sorted(rows_by_sheet.items()):
        # Find matching sheet (case-insensitive)
        ws = None
        for sheet_name in wb.sheetnames:
            if sheet_name.strip().upper() == ticker:
                ws = wb[sheet_name]
                break

        if ws is None:
            warnings_out.append(
                f"[WARN] Ticker {ticker}: no matching sheet — skipped"
            )
            continue

        # Sort all rows for this ticker by acb_date ascending
        rows.sort(key=lambda x: x["acb_date"])

        # Build duplicate key set from current sheet state
        last_data_row = _find_last_data_row(ws)
        existing      = _get_existing_rows(ws, last_data_row)

        # Read reference fonts from the existing sheet ONCE before any insertions
        fonts = _get_ref_fonts(ws, last_data_row)

        for row_data in rows:
            acb_type   = row_data["acb_type"]
            acb_date   = row_data["acb_date"]
            price      = row_data["price"]
            shares     = row_data["shares"]
            commission = row_data["commission"]

            # Find insertion row FIRST (needed for ROC price derivation)
            last_data_row = _find_last_data_row(ws)
            insert_at     = _find_insertion_row(ws, acb_date, last_data_row)
            new_last      = last_data_row + 1

            # Derive ROC per-unit price if not in QIF (needs insert position)
            if acb_type == "ROC" and price is None:
                raw_amount = row_data.get("raw_amount")
                if raw_amount is not None:
                    share_bal = _get_share_balance_before_row(ws, insert_at)
                    if share_bal and share_bal > 0:
                        price = _derive_roc_price(raw_amount, share_bal)
                if price is None:
                    w = (f"[SKIP] {ticker} {acb_date} ROC: cannot derive per-unit price "
                         f"— no share balance found; enter col C manually")
                    warnings_out.append(w)
                    if verbose:
                        print(w)
                    late_skipped += 1
                    continue

            # Single unified dup check — after price is known for all types
            if acb_type in ("Buy", "Sell"):
                price_key  = round(float(price), 2) if price is not None else None
                shares_key = round(float(shares), 4) if shares is not None else None
            else:  # ROC
                price_key  = round(float(price), 5) if price is not None else None
                shares_key = None
            dup_key = (acb_date, acb_type, price_key, shares_key)
            if dup_key in existing:
                w = (f"[SKIP-DUP] {ticker} {acb_date} {acb_type} "
                     f"price={price} — already in sheet")
                warnings_out.append(w)
                if verbose:
                    print(w)
                skipped += 1
                continue

            # Insert a blank row at the target position
            ws.insert_rows(insert_at)
            
            # --- Write input columns A–E with correct formatting ---
            # Col A: Date — store as string in workbook format
            date_str  = _format_date_str(acb_date)
            date_cell = ws.cell(insert_at, 1, date_str)
            _apply_cell(date_cell, DATE_FMT, fonts[1], _COL_ALIGN[1], _BORDER)

            # Col B: Transaction type
            type_cell = ws.cell(insert_at, 2, acb_type)
            _apply_cell(type_cell, "General", fonts[2], _COL_ALIGN[2], _BORDER)

            # Col C: Price / ROC per-unit
            if price is not None:
                price_cell = ws.cell(insert_at, 3, float(price))
                _apply_cell(price_cell, PRICE_FMT, fonts[3], _COL_ALIGN[3], _BORDER)
            else:
                price_cell = ws.cell(insert_at, 3)
                _apply_cell(price_cell, PRICE_FMT, fonts[3], _COL_ALIGN[3], _BORDER)

            # Col D: Shares (blank for ROC)
            if shares is not None:
                shares_cell = ws.cell(insert_at, 4, float(shares))
            else:
                shares_cell = ws.cell(insert_at, 4)
            _apply_cell(shares_cell, SHARES_FMT, fonts[4], _COL_ALIGN[4], _BORDER)

            # Col E: Commission (blank if none)
            if commission is not None:
                comm_cell = ws.cell(insert_at, 5, float(commission))
            else:
                comm_cell = ws.cell(insert_at, 5)
            _apply_cell(comm_cell, COMM_FMT, fonts[5], _COL_ALIGN[5], _BORDER)

            # --- Rewrite ALL formulas from inserted row to new last row ---
            _rewrite_formulas_from(ws, insert_at, new_last, fonts)

            # Update tracking
            existing.add(dup_key)
            sheets_written[ticker] = sheets_written.get(ticker, 0) + 1

            if verbose:
                print(f"[INSERT] {ticker} row {insert_at}: "
                      f"{date_str} {acb_type} price={price} shares={shares}")

        # Quick sanity check: verify formula chain is intact on all formula cols
        final_last = _find_last_data_row(ws)
        for r in range(2, final_last + 1):
            if ws.cell(r, 2).value:
                for col in (6, 7, 8, 9, 10):
                    val = ws.cell(r, col).value
                    if val and not str(val).startswith("="):
                        rollback = True
                        rollback_reason = (
                            f"Formula integrity check failed: "
                            f"sheet {ticker} row {r} col {col}: {val!r}"
                        )
                        break
            if rollback:
                break

        if rollback:
            break

    if rollback:
        if verbose:
            print(f"[ROLLBACK] {rollback_reason}")
        try:
            shutil.copy2(str(backup_path), str(acb_path))
            warnings_out.append(f"[ROLLBACK] Restored from backup. Reason: {rollback_reason}")
        except OSError as exc:
            warnings_out.append(f"[ERROR] Rollback failed: {exc}")
        return {
            "fatal":         rollback_reason,
            "warnings":      warnings_out,
            "imported":      0,
            "skipped":       skipped + late_skipped,
            "rows_by_sheet": rows_by_sheet,
            "backup_path":   str(backup_path),
        }

    try:
        wb.save(str(acb_path))
    except Exception as exc:
        return {"fatal": f"Cannot save workbook: {exc}",
                "warnings": warnings_out, "imported": imported,
                "skipped": skipped, "rows_by_sheet": rows_by_sheet,
                "backup_path": str(backup_path)}

    return {
        "fatal":          None,
        "warnings":       warnings_out,
        "imported":       imported - late_skipped,
        "skipped":        skipped + late_skipped,
        "rows_by_sheet":  rows_by_sheet,
        "sheets_written": sheets_written,
        "backup_path":    str(backup_path),
    }
