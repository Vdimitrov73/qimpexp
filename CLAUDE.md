# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- **Run app**: `python qimpexp.py` (interactive menu) or `python qimpexp.py --acb PATH --year YYYY`
- **Run tests**: `python -m pytest test_qimpexp.py -v` or `python -m unittest test_qimpexp.py -v`
- **Single test class**: `python -m unittest test_qimpexp.TestROCShareFix -v`
- **Single test**: `python -m unittest test_qimpexp.TestROCShareFix.test_phantom_amount_correct -v`
- **Install deps**: `pip install openpyxl`
- **Build exe**: `pyinstaller --onefile --console --name qimpexp --hidden-import openpyxl qimpexp.py`
- **Git tag/release**: `git tag -a v1.0.0 -m "v1.0.0" && git push origin v1.0.0` (triggers CI)
- **CI**: GitHub Actions (`.github/workflows/build_exe.yml`) + GitLab CI (`.gitlab-ci.yml`) — both build .exe + release ZIP on version tags

## Project structure

```
qimpexp.py              CLI entry point, arg parsing, interactive menu
parsers.py              Workbook reading, config loading, account/security resolution
qif_writer.py           QIF building and file writing
qif_importer.py         Phase 2 — QIF → ACB reverse import
ca_calendar.py          Canadian TSX market calendar (T+1 settlement)
qif_colors.py           ANSI colour utilities for interactive menu
ExportToACB.sample.vbs  Sample Quicken toolbar button script (personal copy: ExportToACB.vbs, git-ignored)
test_qimpexp.py         Full test suite (stdlib unittest, no test runner needed)
```

## Architecture

### Two phases

1. **Export** (ACB → QIF): Reads `.xlsx` workbook via `openpyxl`, parses rows per-sheet, reconstructs running share balances, resolves account/security names, builds QIF records grouped by Quicken account, writes one `.qif` file per account.

2. **Import** (QIF → ACB): Parses Quicken-exported QIF files, maps security names back to tickers via `security_map.json`, inserts rows in chronological order with `ws.insert_rows()`, rewrites formulas below insertion point.

### Key design decisions

- **No classes** — functional style with pure functions, dict-based transaction records
- **`Decimal`** for all financial calculations (no float)
- **T+1 settlement** — `ca_calendar.py` handles TSX trading days (Easter via Spencer Jones, stdlib only, no external deps)
- **ROC share balance** — Column G (Share Balance) is authoritative for ROC amount. Falls back to running balance from Buy/Sell rows when col G absent/zero. Col D on ROC rows is ignored with a warning.
- **Duplicate detection** — fuzzy price rounding: Buy/Sell uses 2dp (catches Quicken truncation), ROC uses 5dp (distinguishes same-date distributions)
- **Formula integrity** — importer rewrites all formulas from insertion point to end of sheet, then validates no formula cells were overwritten with values (rolls back from backup if so)

### Transaction flow

Workbook row → `_parse_row_values()` → raw dict → `reconstruct_share_balances()` (per-ticker chronological running balance, resolves ROC amounts) → `apply_date_filter()` → `_filter_by_mode()` → `resolve_account()` / `resolve_security()` → `build_qif_records()` → `write_qif_files()` or `format_qif_dry_run()`

### Config files (JSON, co-located with workbook)

- `account_periods.json` — ticker → list of `{account, start, end}` with overlap validation
- `security_map.json` — ticker → Quicken security name (exact match required for import)

### Warning vs fatal convention

- **Fatal** — returned as `{"fatal": "message"}` string, stops pipeline
- **Warnings** — accumulated in a `warnings` list, printed at end, does not stop processing
