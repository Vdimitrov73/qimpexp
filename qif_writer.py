"""
qif_writer.py - Build and write Quicken QIF output.

Date format: Quicken native M/D'YY with space-padded day (e.g. 9/ 2'25).
Buy/Sell date: settlement date converted back to trade date via prev_trading_day().
RtrnCapX blocks: D N Y U T L[$] lines — no I or Q lines.
Buy/Sell blocks: D N Y I Q U T lines.
"""

from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path

from ca_calendar import prev_trading_day


def _fmt_date(d: date) -> str:
    """Format date in Quicken native M/D'YY format (space-padded day)."""
    return f"{d.month}/{d.day:2d}'{d.year % 100:02d}"


def _fmt_amount(value) -> str:
    """Format dollar amount with 2 decimal places and thousands commas."""
    if value is None:
        return "0.00"
    val = value.quantize(Decimal("0.01"))
    negative = val < 0
    abs_val  = abs(val)
    int_part = int(abs_val)
    frac     = abs_val - int_part
    result   = f"{int_part:,}" + f"{frac:.2f}"[1:]
    return ("-" + result) if negative else result


def _fmt_decimal(value, places: int = 8, min_dp: int = 2) -> str:
    """Format price / quantity — no commas, configurable minimum decimal places.

    min_dp=2  for price (I line) — Quicken always shows e.g. I35.30 not I35.3
    min_dp=0  for shares (Q line) — whole counts stay whole: Q17 not Q17.00
    """
    if value is None:
        return "0"
    q = value.quantize(Decimal("0." + "0" * places))
    s = str(q)
    if "." in s:
        s = s.rstrip("0")
        if s.endswith("."):
            # Whole number after stripping
            s = s.rstrip(".") + ("." + "0" * min_dp if min_dp > 0 else "")
        else:
            integer_part, decimal_part = s.split(".")
            if len(decimal_part) < min_dp:
                decimal_part = decimal_part.ljust(min_dp, "0")
            s = f"{integer_part}.{decimal_part}"
    else:
        s = s + ("." + "0" * min_dp if min_dp > 0 else "")
    return s


def _build_qif_block(txn: dict) -> list:
    """Return list of QIF field lines for one transaction (no trailing ^)."""
    lines  = []
    action = txn["normalized_action"]

    # Buy/Sell: ACB stores settlement date → convert back to trade date for Quicken
    # ROC: ACB stores record date = QIF date → no conversion needed
    if action in ("Buy", "Sell"):
        qif_date = prev_trading_day(txn["trade_date"])
    else:
        qif_date = txn["trade_date"]
    lines.append(f"D{_fmt_date(qif_date)}")
    lines.append(f"N{action}")
    lines.append(f"Y{txn['security_name']}")

    if action == "RtrnCapX":
        amount = txn.get("total_amount")
        if amount is not None:
            lines.append(f"U{_fmt_amount(amount)}")
            lines.append(f"T{_fmt_amount(amount)}")
            lines.append(f"L[{txn.get('account_name', '')}]")
            lines.append(f"${_fmt_amount(amount)}")
    else:
        price  = txn.get("price_per_share")
        shares = txn.get("shares")
        amount = txn.get("total_amount")
        if price  is not None: lines.append(f"I{_fmt_decimal(price,  min_dp=2)}")
        if shares is not None: lines.append(f"Q{_fmt_decimal(shares, min_dp=0)}")
        if amount is not None:
            lines.append(f"U{_fmt_amount(amount)}")
            lines.append(f"T{_fmt_amount(amount)}")

    return lines


def _qif_sort_date(txn: dict) -> date:
    if txn["normalized_action"] in ("Buy", "Sell"):
        return prev_trading_day(txn["trade_date"])
    return txn["trade_date"]


_ACTION_PRIORITY = {"Sell": 0, "Buy": 1, "RtrnCapX": 2}


def build_qif_records(transactions: list, verbose: bool = False) -> dict:
    by_account = defaultdict(list)
    for txn in transactions:
        by_account[txn["account_name"]].append(txn)

    for account in by_account:
        by_account[account].sort(
            key=lambda t: (
                _qif_sort_date(t),
                _ACTION_PRIORITY.get(t["normalized_action"], 9),
                t.get("ticker", ""),            # alphabetical within same date+action
                t["source_row_number"] or 0,    # within same ticker, preserve sheet row order
            )
        )
        
    qif_groups = {}
    for account, txns in sorted(by_account.items()):
        blocks = []
        for txn in txns:
            block = _build_qif_block(txn)
            blocks.append(block)
            if verbose:
                print(f"[QIF] {account} | {txn['trade_date']} {txn['normalized_action']} "
                      f"{txn['security_name']} "
                      f"shares={txn.get('shares')} price={txn.get('price_per_share')} "
                      f"amount={txn.get('total_amount')}")
        qif_groups[account] = blocks

    return qif_groups


def _render_qif(account_name: str, blocks: list, start_date=None) -> str:
    lines = ["!Type:Invst"]
    if start_date is not None:
        lines.append(f"D{_fmt_date(start_date)}")
        lines.append("NCash")
        lines.append(f"L[{account_name}]")
        lines.append("^")
    for block in blocks:
        lines.extend(block)
        lines.append("^")
    return "\n".join(lines) + "\n"


def _build_filename(account_name: str, start_date, end_date) -> str:
    safe      = account_name.replace(" ", "_").replace("/", "-")
    start_str = start_date.strftime("%Y%m%d") if start_date else "unknown"
    end_str   = end_date.strftime("%Y%m%d")   if end_date   else "unknown"
    return f"{safe}_{start_str}_{end_str}.qif"


def write_qif_files(qif_groups: dict, output_dir, start_date, end_date,
                    verbose: bool = False) -> list:
    output_dir = Path(output_dir)
    written    = []
    for account_name, blocks in sorted(qif_groups.items()):
        filename = _build_filename(account_name, start_date, end_date)
        out_path = output_dir / filename
        content  = _render_qif(account_name, blocks, start_date=start_date)
        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        if verbose:
            print(f"[WRITE] {out_path} ({len(blocks)} transactions)")
        written.append(out_path)
    return written


def format_qif_dry_run(qif_groups: dict, start_date=None) -> str:
    sections = []
    for account_name, blocks in sorted(qif_groups.items()):
        header  = (f"\n{'='*50}\nACCOUNT: {account_name}  "
                   f"({len(blocks)} transactions)\n{'='*50}")
        content = _render_qif(account_name, blocks, start_date=start_date)
        sections.append(header + "\n" + content)
    return "\n".join(sections) if sections else "(no transactions)"
