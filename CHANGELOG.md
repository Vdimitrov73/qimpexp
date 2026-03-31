# Changelog

All notable changes to QImpExp are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Versioning follows [Semantic Versioning](https://semver.org/).

---

## [1.0.0] — 2025-03-30

### Added
- **Phase 1 — ACB → QIF export**
  - Reads `.xlsx` ACB workbook via `openpyxl` (no external parser dependencies for core logic)
  - Supports Buy, Sell, ROC (return of capital), and phantom distribution transactions
  - Multi-account support via `account_periods.json` — maps each ticker to one or more
    Quicken account names with date ranges
  - Security name mapping via `security_map.json`
  - Date filters: `--year YYYY`, `--start`, `--end`
  - Mode filters: `full`, `tax-adjustments` (ROC only), `buys-sells`
  - `--dry-run` preview without writing files
  - `--verbose` detailed parse/mapping trace
  - One QIF file per Quicken account, filename `<Account>_<YYYYMMDD>_<YYYYMMDD>.qif`
  - QIF date format: `MM/DD/YYYY`
  - `RtrnCapX` blocks include `L[Account]` and `$amount` lines matching Quicken's native format
  - Amounts formatted with thousands commas (`1,038.48`)

- **ROC share balance logic**
  - Column G (Share Balance) from the workbook is used as the authoritative share count
    for ROC amount computation — never column D
  - Column D on ROC rows is ignored with a warning if populated
  - Falls back to running balance reconstructed from Buy/Sell rows when column G is absent
  - Share balance reconstruction happens before date filtering, so mid-history date ranges
    always have the correct balance
  - Warning emitted when share balance goes negative after a Sell
  - Zero-share ROC rows are skipped with a warning

- **Phase 2 — QIF → ACB import**
  - Reads Quicken-exported QIF files (both native `M/D'YY` and standard `MM/DD/YYYY` dates)
  - Supports Buy, Sell, RtrnCapX; other actions skipped with a note
  - Buy/Sell: Quicken trade date → settlement date (T+1 Canadian trading day) written to ACB
  - RtrnCapX: record date unchanged (same in both Quicken and ACB)
  - Appends rows to the correct sheet with ACB formula strings in columns F–J
  - `--dry-run` preview without touching the workbook

- **Canadian TSX market calendar** (`ca_calendar.py`)
  - Pure stdlib (no external deps): `next_trading_day()`, `prev_trading_day()`,
    `is_trading_day()`, `canadian_holidays()`
  - Covers: New Year, Family Day, Good Friday, Victoria Day, Canada Day, Civic Holiday,
    Labour Day, Thanksgiving, Christmas, Boxing Day
  - Easter computed via Spencer Jones algorithm

- **Interactive menu mode** (`qimpexp.py` with no arguments)
  - First-time setup wizard: security names + account periods, with save to JSON
  - Main menu: export, export preview, import, import preview, edit config, show config,
    change workbook
  - ANSI colour UI via `qif_colors.py`; graceful fallback when colours unavailable

- **`qif_colors.py`** — shared ANSI colour utilities, uses `col()` function

- **CI/CD pipelines**: GitHub Actions (`build_exe.yml`) and GitLab CI (`.gitlab-ci.yml`)
  — triggered on version tags, builds `qimpexp.exe` on Windows, creates release ZIP
  with exe + `README.FIRST.txt` + sample data
