"""
QImpExp - ACB Excel workbook <-> Quicken QIF converter.

Modes:
  CLI:         qimpexp.py --acb FILE [options]
  Interactive: qimpexp.py   (no arguments)
"""

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Colour support — try project colours; silent fallback if absent
# ---------------------------------------------------------------------------
try:
    from qif_colors import use_color, col, BOLD, DIM, CYAN, GREEN, YELLOW, RED, RESET
except ImportError:
    try:
        from t3_colors import use_color, col, BOLD, DIM, CYAN, GREEN, YELLOW, RED, RESET
    except ImportError:
        def use_color(): return False
        def col(code, text): return text
        BOLD = DIM = CYAN = GREEN = YELLOW = RED = RESET = ""

from parsers import (
    load_workbook_transactions,
    load_account_periods,
    save_account_periods,
    load_security_map,
    save_security_map,
    interactive_account_setup,
    interactive_security_setup,
    resolve_account,
    resolve_security,
    apply_date_filter,
    reconstruct_share_balances,
)
from qif_writer  import build_qif_records, write_qif_files, format_qif_dry_run
from qif_importer import import_qif_to_acb


# ===========================================================================
# Coloured UI helpers
# ===========================================================================

_W = 65   # menu box inner width (between pipes, including spaces)


def _box_top():
    print(col(BOLD + YELLOW, "╔" + "═" * _W + "╗"))


def _box_mid():
    print(col(BOLD + YELLOW, "║" + "═" * _W + "║"))


def _box_bot():
    print(col(BOLD + YELLOW, "╚" + "═" * _W + "╝"))


def _box_row(text: str):
    # Pad / truncate to _W characters
    padded = f" {text} ".ljust(_W)[:_W]
    print(col(BOLD + YELLOW, "║") + col(BOLD + CYAN, padded) + col(BOLD + YELLOW, "║"))


def _box_title(text: str):
    padded = text.center(_W)[:_W]
    print(col(BOLD + YELLOW, "║") + col(BOLD + YELLOW, padded) + col(BOLD + YELLOW, "║"))


def _print_menu():
    print()
    _box_top()
    _box_title("  QImpExp  —  ACB / Quicken QIF Converter  ")
    _box_mid()
    _box_row("  1  Export ACB -> QIF          Generate QIF from workbook")
    _box_row("  2  Preview export (dry run)   Show QIF without writing")
    _box_row("  3  Import QIF -> ACB          Append QIF rows to workbook")
    _box_row("  4  Preview import (dry run)   Show rows without writing")
    _box_mid()
    _box_row("  5  Edit account periods        Manage ticker <-> account map")
    _box_row("  6  Show loaded config          Display current settings")
    _box_row("  7  Change workbook             Switch to a different .xlsx")
    _box_mid()
    _box_row("  Q  Quit")
    _box_bot()
    print()


def _ok(text: str):
    print(col(GREEN, "  ✓ " + text))


def _warn(text: str):
    print(col(YELLOW, "  ! " + text))


def _err(text: str):
    print(col(RED, "  ✗ " + text))


def _head(text: str):
    print()
    print(col(BOLD + YELLOW, "═" * (_W + 2)))
    print(col(BOLD + YELLOW, f"  {text}"))
    print(col(BOLD + YELLOW, "═" * (_W + 2)))


# ===========================================================================
# CLI argument parsing
# ===========================================================================

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog="qimpexp",
        description="Convert ACB Excel workbook to Quicken QIF files, and back.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Modes:
  full              Buy + Sell + ROC  (default)
  tax-adjustments   ROC only
  buys-sells        Buy and Sell only

Date filtering:
  --year 2024       equivalent to --start 2024-01-01 --end 2024-12-31

Config files (same directory as workbook):
  account_periods.json    ticker -> Quicken account name(s) with date ranges
  security_map.json       ticker -> Quicken security name

Run without arguments for interactive menu mode.
        """,
    )
    parser.add_argument("--acb",     metavar="PATH", help="Path to ACB .xlsx workbook")
    parser.add_argument("--account", metavar="NAME", help="Default account name")
    parser.add_argument("--funds",   nargs="+", metavar="FUND")
    parser.add_argument("--start",   metavar="YYYY-MM-DD")
    parser.add_argument("--end",     metavar="YYYY-MM-DD")
    parser.add_argument("--year",    metavar="YYYY")
    parser.add_argument("--output",  metavar="PATH", default=".")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--mode",    choices=["full", "tax-adjustments", "buys-sells"],
                        default="full")
    # Phase 2: import
    parser.add_argument("--import-qif", metavar="QIF_PATH",
                        help="Import a QIF file into the ACB workbook (Phase 2)")
    return parser.parse_args(argv)


def _parse_date_arg(value, name):
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        print(f"ERROR: Invalid date for {name}: {value!r}", file=sys.stderr)
        sys.exit(1)


def _resolve_date_range(args):
    start = end = None
    if args.year:
        year  = int(args.year)
        start = date(year, 1, 1)
        end   = date(year, 12, 31)
    if args.start:
        start = _parse_date_arg(args.start, "--start")
    if args.end:
        end   = _parse_date_arg(args.end,   "--end")
    return start, end


# ===========================================================================
# Core pipeline (shared by CLI and menu)
# ===========================================================================

def _load_config(acb_path, verbose):
    acb_dir = Path(acb_path).parent
    ap_path = acb_dir / "account_periods.json"
    sm_path = acb_dir / "security_map.json"

    account_periods = security_map = None

    if ap_path.exists():
        try:
            account_periods = load_account_periods(ap_path)
            if verbose:
                print(f"[CONFIG] account_periods.json ({len(account_periods)} tickers)")
        except ValueError as exc:
            print(f"FATAL: {exc}", file=sys.stderr); sys.exit(1)

    if sm_path.exists():
        try:
            security_map = load_security_map(sm_path)
            if verbose:
                print(f"[CONFIG] security_map.json ({len(security_map)} entries)")
        except ValueError as exc:
            print(f"FATAL: {exc}", file=sys.stderr); sys.exit(1)

    return account_periods, security_map, ap_path, sm_path


def _filter_by_mode(transactions, mode):
    if mode == "full":            return transactions
    if mode == "tax-adjustments": return [t for t in transactions if t["normalized_action"] == "RtrnCapX"]
    if mode == "buys-sells":      return [t for t in transactions if t["normalized_action"] in ("Buy","Sell")]
    return transactions


def run_export_pipeline(
    acb_path, fund_filter=None, start_date=None, end_date=None,
    mode="full", output_dir=".", dry_run=False, verbose=False,
    default_account=None, account_periods=None, security_map=None,
):
    """ACB workbook -> QIF.  Returns summary dict."""
    warnings = []

    parse_result = load_workbook_transactions(acb_path, fund_filter=fund_filter, verbose=verbose)
    if parse_result["fatal"]:
        return {"fatal": parse_result["fatal"]}

    all_txns = parse_result["transactions"]
    warnings.extend(parse_result["warnings"])

    if not all_txns:
        return {"fatal": "No valid transactions found."}

    # reconstruct_share_balances now returns (txns, extra_warnings)
    all_txns, bal_warns = reconstruct_share_balances(all_txns, verbose=verbose)
    warnings.extend(bal_warns)

    filtered     = apply_date_filter(all_txns, start_date, end_date, verbose=verbose)
    if not filtered:
        return {"fatal": "No transactions remain after date filtering."}

    mode_filtered = _filter_by_mode(filtered, mode)
    tickers_needed = sorted({t["ticker"] for t in mode_filtered})

    if account_periods is None and default_account is None:
        print("\nNo account_periods.json found and --account not specified.")
        account_periods = interactive_account_setup(tickers_needed)
    if security_map is None:
        print("\nNo security_map.json found.")
        security_map = interactive_security_setup(tickers_needed)

    resolved = []
    unresolved_tickers = set()

    for txn in mode_filtered:
        account = resolve_account(txn, account_periods, default_account=default_account,
                                  verbose=verbose)
        if account is None:
            unresolved_tickers.add(txn["ticker"])
            continue
        security = resolve_security(txn, security_map, verbose=verbose)
        txn = dict(txn)
        txn["account_name"] = account
        txn["security_name"] = security
        resolved.append(txn)

    if account_periods is not None and unresolved_tickers:
        print(f"\n[WARN] These tickers have no account mapping: "
              f"{', '.join(sorted(unresolved_tickers))}")
        answer = input("Run setup wizard for missing tickers? (y/n): ").strip().lower()
        if answer == "y":
            new_periods = interactive_account_setup(sorted(unresolved_tickers))
            account_periods.update(new_periods)
            # Re-resolve only the previously unresolved transactions
            for txn in mode_filtered:
                if txn["ticker"] not in unresolved_tickers:
                    continue
                account = resolve_account(txn, account_periods,
                                          default_account=default_account,
                                          verbose=verbose)
                if account is None:
                    continue
                security = resolve_security(txn, security_map, verbose=verbose)
                txn = dict(txn)
                txn["account_name"] = account
                txn["security_name"] = security
                resolved.append(txn)

    for txn in mode_filtered:
        if txn["ticker"] in unresolved_tickers and not any(
                r["ticker"] == txn["ticker"] for r in resolved):
            w = (f"[SKIP] {txn['source_sheet']} row {txn['source_row_number']}: "
                 f"no account for {txn['ticker']} on {txn['trade_date']} — skipped")
            if w not in warnings:
                warnings.append(w)

    skipped = len(mode_filtered) - len(resolved)

    if not resolved:
        return {"fatal": "No transactions could be resolved to an account."}

    qif_groups = build_qif_records(resolved, verbose=verbose)
    all_dates  = [t["trade_date"] for t in resolved]
    disp_start = min(all_dates)
    disp_end   = max(all_dates)

    output_files = []
    if dry_run:
        print(format_qif_dry_run(qif_groups, start_date=start_date or disp_start))
    else:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        output_files = write_qif_files(
            qif_groups, output_dir=out,
            start_date=start_date or disp_start,
            end_date=end_date   or disp_end,
            verbose=verbose,
        )

    return {
        "fatal": None, "warnings": warnings,
        "all_transactions": all_txns, "filtered": filtered,
        "mode_filtered": mode_filtered, "resolved": resolved,
        "skipped": skipped, "output_files": output_files,
        "tickers": tickers_needed,
        "disp_start": disp_start, "disp_end": disp_end,
        "dry_run": dry_run,
    }


def _print_export_summary(result, acb_path, mode, start_date, end_date):
    warnings = result.get("warnings", [])
    if warnings:
        print()
        print(col(YELLOW, "--- WARNINGS ---"))
        for w in warnings:
            print(col(YELLOW, w))

    print()
    print(col(BOLD + CYAN, "=" * 44))
    print(col(BOLD + CYAN, "  QUICKEN IMPORTER SUMMARY"))
    print(col(BOLD + CYAN, "=" * 44))
    print(f"  Workbook:   {acb_path}")
    tickers = result.get("tickers", [])
    print(f"  Funds:      {', '.join(tickers) if tickers else '(all)'}")
    s = start_date or result.get("disp_start", "?")
    e = end_date   or result.get("disp_end",   "?")
    print(f"  Date range: {s} to {e}")
    print(f"  Mode:       {mode}")
    print()
    print(f"  Read:       {len(result.get('all_transactions', []))}")
    print(f"  After date: {len(result.get('filtered', []))}")
    print(f"  After mode: {len(result.get('mode_filtered', []))}")
    print(f"  Written:    {col(GREEN, str(len(result.get('resolved', []))))}")
    print(f"  Skipped:    {col(YELLOW, str(result.get('skipped', 0)))}")
    print(f"  Warnings:   {col(YELLOW, str(len(warnings)))}")

    output_files = result.get("output_files", [])
    if output_files:
        print()
        print(col(BOLD, "  Output files:"))
        for f in output_files:
            print(col(GREEN, f"    {f}"))
    elif result.get("dry_run"):
        print(col(DIM, "\n  (dry-run: no files written)"))


def _print_import_summary(result, qif_path):
    warnings = result.get("warnings", [])
    if warnings:
        print()
        print(col(YELLOW, "--- WARNINGS ---"))
        for w in warnings:
            print(col(YELLOW, w))

    print()
    print(col(BOLD + CYAN, "=" * 44))
    print(col(BOLD + CYAN, "  QIF IMPORT SUMMARY"))
    print(col(BOLD + CYAN, "=" * 44))
    print(f"  QIF file:   {qif_path}")
    print()
    print(f"  Imported:   {col(GREEN, str(result.get('imported', 0)))}")
    print(f"  Skipped:    {col(YELLOW, str(result.get('skipped', 0)))}")
    print(f"  Warnings:   {col(YELLOW, str(len(warnings)))}")

    backup = result.get("backup_path")
    if backup:
        print(col(GREEN, f"  Backup:     {backup}"))

    sheets = result.get("sheets_written", {})
    if sheets:
        print()
        print(col(BOLD, "  Rows appended per sheet:"))
        for ticker, count in sorted(sheets.items()):
            print(col(GREEN, f"    {ticker}: {count}"))


# ===========================================================================
# Interactive menu helpers
# ===========================================================================

def _input_required(prompt):
    while True:
        val = input(col(CYAN, prompt)).strip()
        if val: return val
        _warn("Required — cannot be empty.")


def _input_optional(prompt, default=""):
    val = input(col(CYAN, prompt)).strip()
    return val if val else default


def _input_date(prompt, allow_empty=True):
    while True:
        raw = input(col(CYAN, prompt)).strip()
        if not raw and allow_empty: return None
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            _warn("Use YYYY-MM-DD format.")


def _input_choice(prompt, options, allow_empty=False):
    """options: list of (key, label).  Returns matched key (lowercased)."""
    while True:
        raw = input(col(CYAN, prompt)).strip().lower()
        if not raw and allow_empty: return None
        for key, _ in options:
            if raw == key.lower(): return key
        _warn(f"Enter one of: {'/'.join(k for k, _ in options)}")


# ===========================================================================
# First-time setup wizard
# ===========================================================================

def _wizard_new_config(acb_path, tickers_found):
    _head("FIRST-TIME SETUP WIZARD")
    print(f"  Found {len(tickers_found)} fund(s): {col(CYAN, ', '.join(tickers_found))}")
    print()

    acb_dir = Path(acb_path).parent

    # Step 1 – security names
    print(col(BOLD + YELLOW, "  Step 1 of 2 — Security Names"))
    print(col(DIM, "  Enter the exact Quicken security name (Enter = use ticker).\n"))
    security_map = {}
    for ticker in sorted(tickers_found):
        name = input(col(CYAN, f"    {ticker} security name [{ticker}]: ")).strip()
        security_map[ticker] = name if name else ticker

    # Step 2 – account periods
    print()
    print(col(BOLD + YELLOW, "  Step 2 of 2 — Account Mapping"))
    print(col(DIM, "  Map each ticker to the Quicken account where it is held.\n"))

    account_periods = {}
    for ticker in sorted(tickers_found):
        print(f"  {col(CYAN, ticker)}")
        periods = []
        while True:
            account  = _input_required(f"    Account name: ")
            start_dt = _input_date("    Start date (YYYY-MM-DD, blank=beginning): ", True) or date(1900, 1, 1)
            end_dt   = _input_date("    End date   (YYYY-MM-DD, blank=still active): ", True)
            periods.append({"account": account, "start": start_dt, "end": end_dt})
            more = _input_choice(f"    Add another period for {ticker}? (y/n): ",
                                 [("y","yes"),("n","no")])
            if more == "n": break
        account_periods[ticker] = periods

    # Offer to save
    print()
    save = _input_choice("  Save config to account_periods.json / security_map.json? (y/n): ",
                         [("y","yes"),("n","no")])
    if save == "y":
        ap_path = acb_dir / "account_periods.json"
        sm_path = acb_dir / "security_map.json"
        save_account_periods(ap_path, account_periods)
        save_security_map(sm_path, security_map)
        _ok(f"Saved {ap_path}")
        _ok(f"Saved {sm_path}")

    return account_periods, security_map


# ===========================================================================
# Menu: edit account periods
# ===========================================================================

def _menu_edit_config(account_periods, security_map, acb_dir):
    _head("EDIT CONFIG")

    if account_periods:
        print(col(BOLD, "  Account periods:"))
        for ticker in sorted(account_periods):
            for p in account_periods[ticker]:
                end_str = p["end"].strftime("%Y-%m-%d") if p["end"] else "open"
                print(f"    {col(CYAN, ticker):12s} | {p['account']:30s} | {p['start']} – {end_str}")
    else:
        _warn("No account periods loaded.")

    print()
    if security_map:
        print(col(BOLD, "  Security map:"))
        for k, v in sorted(security_map.items()):
            print(f"    {col(CYAN, k):12s} -> {v}")
    else:
        _warn("No security map loaded.")

    print()
    action = _input_choice(
        "  (a)dd ticker  (e)dit period  (s)ave  (q)uit: ",
        [("a","add"), ("e","edit"), ("s","save"), ("q","quit")],
    )

    if action == "a":
        ticker = _input_required("  Ticker (uppercase): ").upper()
        name   = input(col(CYAN, f"  Security name for {ticker} [{ticker}]: ")).strip()
        security_map[ticker] = name if name else ticker
        periods = []
        while True:
            account  = _input_required("  Account: ")
            start_dt = _input_date("  Start (YYYY-MM-DD, blank=beginning): ", True) or date(1900,1,1)
            end_dt   = _input_date("  End   (YYYY-MM-DD, blank=still active): ", True)
            periods.append({"account": account, "start": start_dt, "end": end_dt})
            more = _input_choice(f"  Add another period? (y/n): ", [("y","yes"),("n","no")])
            if more == "n": break
        account_periods[ticker] = periods
        _offer_save(account_periods, security_map, acb_dir)

    elif action == "e":
        ticker = _input_required("  Ticker to edit: ").upper()
        if ticker not in account_periods:
            _warn(f"{ticker} not found."); return
        for i, p in enumerate(account_periods[ticker]):
            print(f"  [{i}] {p['account']}  {p['start']} – {p['end'] or 'open'}")
        idx = int(_input_required("  Index to edit: "))
        p = account_periods[ticker][idx]
        new_account = input(col(CYAN, f"  Account [{p['account']}]: ")).strip() or p["account"]
        new_start   = _input_date(f"  Start [{p['start']}]: ", True) or p["start"]
        new_end     = _input_date(f"  End [{p['end'] or 'open'}]: ", True)
        account_periods[ticker][idx] = {"account": new_account, "start": new_start, "end": new_end}
        _offer_save(account_periods, security_map, acb_dir)

    elif action == "s":
        _offer_save(account_periods, security_map, acb_dir)


def _offer_save(account_periods, security_map, acb_dir):
    save = _input_choice("  Save changes? (y/n): ", [("y","yes"),("n","no")])
    if save == "y":
        ap_path = acb_dir / "account_periods.json"
        sm_path = acb_dir / "security_map.json"
        save_account_periods(ap_path, account_periods)
        save_security_map(sm_path, security_map)
        _ok(f"Saved {ap_path}")
        _ok(f"Saved {sm_path}")


# ===========================================================================
# Interactive export options
# ===========================================================================

def _collect_export_options(dry_run: bool):
    print()
    _box_top()
    _box_title("EXPORT OPTIONS — " + ("DRY RUN PREVIEW" if dry_run else "WRITE QIF FILES"))
    _box_mid()
    _box_row("  Mode:")
    _box_row("    1  full             Buy + Sell + ROC")
    _box_row("    2  tax-adjustments  ROC only")
    _box_row("    3  buys-sells       Buy + Sell only")
    _box_mid()
    _box_row("  Date range:")
    _box_row("    Enter a year  OR  separate start / end dates")
    _box_mid()
    _box_row("  Funds:")
    _box_row("    Space-separated tickers, blank = all")
    if not dry_run:
        _box_mid()
        _box_row("  Output directory:")
        _box_row("    Blank = current directory")
    _box_mid()
    _box_row("  M  Back to main menu")
    _box_bot()
    print()

    mc = _input_choice("  Mode [1]: ",
                       [("1","f"),("2","t"),("3","b"),("m","menu")],
                       allow_empty=True) or "1"
    if mc == "m":
        return None                          # ← signal: go back

    mode = {"1":"full","2":"tax-adjustments","3":"buys-sells"}[mc]

    year_raw = input(col(CYAN, "  Year (e.g. 2025, blank to enter dates): ")).strip()
    if year_raw:
        try:
            yr = int(year_raw)
            start_date, end_date = date(yr, 1, 1), date(yr, 12, 31)
        except ValueError:
            start_date = end_date = None
    else:
        start_date = _input_date("  Start date (YYYY-MM-DD, blank=no limit): ", True)
        end_date   = _input_date("  End date   (YYYY-MM-DD, blank=no limit): ", True)

    funds_raw   = input(col(CYAN, "  Funds (space-separated, blank=all): ")).strip()
    fund_filter = funds_raw.upper().split() if funds_raw else None

    if dry_run:
        output_dir = "."
    else:
        out_raw    = input(col(CYAN, "  Output directory [current dir]: ")).strip()
        output_dir = out_raw if out_raw else "."

    verbose_ch = _input_choice("  Verbose? (y/n) [n]: ",
                               [("y","yes"),("n","no")], allow_empty=True) or "n"
    verbose = (verbose_ch == "y")

    return mode, start_date, end_date, fund_filter, output_dir, verbose


# ===========================================================================
# Interactive import options
# ===========================================================================

def _collect_import_options(dry_run: bool, acb_path):
    _head("IMPORT OPTIONS — " + ("DRY RUN PREVIEW" if dry_run else "WRITE TO WORKBOOK"))
    print(col(DIM, "  Type M at any prompt to return to the main menu.\n"))

    while True:
        raw = _input_required("  QIF file path (M = main menu): ")
        if raw.strip().lower() == "m":
            return None                      # ← signal: go back
        p = Path(raw.strip('"').strip("'"))
        if p.exists():
            qif_path = p
            break
        _err(f"File not found: {p}")

    # Mode selection — mirrors the export mode menu
    print()
    print(col(BOLD, "  Mode:"))
    _box_top()
    _box_row("  1  full             Buy + Sell + ROC")
    _box_row("  2  tax-adjustments  ROC only")
    _box_row("  3  buys-sells       Buy + Sell only")
    _box_bot()
    mc = _input_choice("  Mode [1]: ", [("1","f"),("2","t"),("3","b")],
                       allow_empty=True) or "1"
    mode = {"1": "full", "2": "tax-adjustments", "3": "buys-sells"}[mc]

    verbose_ch = _input_choice("  Verbose? (y/n) [n]: ",
                               [("y","yes"),("n","no")], allow_empty=True) or "n"
    verbose = (verbose_ch == "y")

    return qif_path, mode, verbose


# ===========================================================================
# Interactive main loop
# ===========================================================================

def run_interactive():
    print()
    print(col(BOLD + CYAN, "  QImpExp"))

    # Locate workbook
    acb_path = None
    while acb_path is None:
        raw = _input_required("ACB workbook path (.xlsx): ")
        p   = Path(raw.strip('"').strip("'"))
        if p.exists():
            acb_path = p
        else:
            _err(f"File not found: {p}")

    acb_dir = acb_path.parent

    # Scan sheets
    print(f"\n  Scanning {acb_path.name} ...")
    probe = load_workbook_transactions(acb_path, verbose=False)
    if probe["fatal"]:
        _err(probe["fatal"]); return

    tickers_found = sorted({t["ticker"] for t in probe["transactions"]})
    _ok(f"Found: {', '.join(tickers_found)}")

    # Load or create config
    ap_path = acb_dir / "account_periods.json"
    sm_path = acb_dir / "security_map.json"

    account_periods = security_map = None

    if ap_path.exists() and sm_path.exists():
        try:
            account_periods = load_account_periods(ap_path)
            security_map    = load_security_map(sm_path)
            _ok(f"Loaded config ({len(account_periods)} ticker(s))")
        except ValueError as exc:
            _warn(f"Config error: {exc}")
            account_periods, security_map = _wizard_new_config(acb_path, tickers_found)
    else:
        missing = [n for n, p in [("account_periods.json", ap_path),
                                   ("security_map.json",    sm_path)] if not p.exists()]
        _warn(f"Missing: {', '.join(missing)}")
        if _input_choice("  Run setup wizard? (y/n): ", [("y","yes"),("n","no")]) == "y":
            account_periods, security_map = _wizard_new_config(acb_path, tickers_found)

    # Main loop
    while True:
        _print_menu()
        choice = _input_choice("Choice: ",
                               [("1",""), ("2",""), ("3",""), ("4",""),
                                ("5",""), ("6",""), ("7",""), ("q","")])

        if choice == "q":
            print(col(DIM, "\nGoodbye.\n")); break

        elif choice == "7":
            raw = _input_required("New workbook path: ")
            p   = Path(raw.strip('"').strip("'"))
            if p.exists():
                acb_path = p; acb_dir = p.parent
                ap_path  = acb_dir / "account_periods.json"
                sm_path  = acb_dir / "security_map.json"
                # Reset config for new workbook
                account_periods = security_map = None
                _ok(f"Workbook: {acb_path}")
                if ap_path.exists():
                    account_periods = load_account_periods(ap_path)
                if sm_path.exists():
                    security_map = load_security_map(sm_path)
            else:
                _err(f"Not found: {p}")

        elif choice == "6":
            _head("LOADED CONFIG")
            if account_periods:
                print(col(BOLD, "  Account periods:"))
                for ticker in sorted(account_periods):
                    for p in account_periods[ticker]:
                        end_str = p["end"].strftime("%Y-%m-%d") if p["end"] else "open"
                        print(f"    {col(CYAN, ticker):12s} | {p['account']:30s} | {p['start']} – {end_str}")
            else:
                _warn("No account periods.")
            print()
            if security_map:
                print(col(BOLD, "  Security map:"))
                for k, v in sorted(security_map.items()):
                    print(f"    {col(CYAN, k):12s} -> {v}")
            else:
                _warn("No security map.")

        elif choice == "5":
            _menu_edit_config(account_periods or {}, security_map or {}, acb_dir)
            if ap_path.exists():
                account_periods = load_account_periods(ap_path)
            if sm_path.exists():
                security_map = load_security_map(sm_path)

        elif choice in ("1", "2"):
            dry_run = (choice == "2")
            result = _collect_export_options(dry_run)
            if result is None:
                continue
            mode, start_date, end_date, fund_filter, output_dir, verbose = result
            print()
            result = run_export_pipeline(
                acb_path=acb_path, fund_filter=fund_filter,
                start_date=start_date, end_date=end_date,
                mode=mode, output_dir=output_dir, dry_run=dry_run,
                verbose=verbose, account_periods=account_periods,
                security_map=security_map,
            )
            if result.get("fatal"):
                _err(result["fatal"])
            else:
                _print_export_summary(result, acb_path, mode, start_date, end_date)

        elif choice in ("3", "4"):
            dry_run = (choice == "4")
            if security_map is None:
                _err("security_map.json is required for import. Set it up first (option 5).")
                continue
            result = _collect_import_options(dry_run, acb_path)
            if result is None:
                continue
            qif_path, mode, verbose = result
            print()
            result = import_qif_to_acb(
                qif_path=qif_path, acb_path=acb_path,
                security_map=security_map,
                mode=mode,
                dry_run=dry_run, verbose=verbose,
            )
            if result.get("fatal"):
                _err(result["fatal"])
            else:
                _print_import_summary(result, qif_path)


# ===========================================================================
# Entry point
# ===========================================================================

def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    if argv:
        args = parse_args(argv)

        if not args.acb:
            parse_args(["--help"]); return 0

        acb_path = Path(args.acb)
        if not acb_path.exists():
            print(f"ERROR: Workbook not found: {acb_path}", file=sys.stderr); sys.exit(1)

        account_periods, security_map, _, _ = _load_config(acb_path, args.verbose)

        # Phase 2: import mode
        if args.import_qif:
            if security_map is None:
                print("ERROR: security_map.json required for import.", file=sys.stderr)
                sys.exit(1)
            result = import_qif_to_acb(
                qif_path=Path(args.import_qif), acb_path=acb_path,
                security_map=security_map,
                mode=args.mode,
                dry_run=args.dry_run, verbose=args.verbose,
            )
            if result.get("fatal"):
                print(f"FATAL: {result['fatal']}", file=sys.stderr); sys.exit(1)
            _print_import_summary(result, args.import_qif)
            return 0

        # Phase 1: export mode
        start_date, end_date = _resolve_date_range(args)
        result = run_export_pipeline(
            acb_path=acb_path, fund_filter=args.funds,
            start_date=start_date, end_date=end_date,
            mode=args.mode, output_dir=args.output,
            dry_run=args.dry_run, verbose=args.verbose,
            default_account=args.account,
            account_periods=account_periods, security_map=security_map,
        )
        if result.get("fatal"):
            print(f"FATAL: {result['fatal']}", file=sys.stderr); sys.exit(1)
        _print_export_summary(result, acb_path, args.mode, start_date, end_date)

    else:
        run_interactive()

    return 0


if __name__ == "__main__":
    sys.exit(main())
