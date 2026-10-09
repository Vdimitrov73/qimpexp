# Changelog

All notable changes to QImpExp are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Versioning follows [Semantic Versioning](https://semver.org/).

---
## [1.1.5] — 2026-10-09

### Added
- Memo-driven DRIP reinvestments: a `Buy` whose memo begins with `DRIP`
  (case-insensitive, e.g. `Drip VRE`) imports at the settlement date with
  no T+1 shift, is marked in column K, and re-exports as `NBuy` + `MDRIP`.
- Import validates the QIF file account against `account_periods.json` for
  each trade date; wrong-account rows skip with a warning (the account is
  now also read from the file's Cash preamble block).
- Exported QIF files start with a Quicken-verified `!Account` header
  identifying the destination account.

### Fixed
- Removed the Cash pseudo-transaction from QIF output — Quicken imported
  it as a zero-amount MiscExp.
- Config editor hardened: guarded period index, blank-means-keep end date,
  and pre-save validation so invalid JSON can no longer be written
  (previously it saved, then crashed the menu on reload).
- Unknown `--mode` now fails fast (`ValueError` on export, fatal on
  import) instead of silently processing everything.
- QIF and JSON config writes are atomic (temp file + replace); Excel
  serial dates corrected by one day; end-before-start periods rejected;
  missing column-G cache and `--year` combined with `--start/--end`
  now warn instead of staying silent.

## [1.1.4] — 2026-10-08

### Fixed
- Non-interactive (scheduled/CLI) exports no longer block on setup-wizard
  prompts: missing config or a mode matching nothing now returns a clean
  fatal instead of hanging on `input()`.
- `--year` with a non-numeric or out-of-range value now exits(1) with an
  error instead of a traceback; a mistyped year in the interactive menu
  warns instead of silently dropping the date filter.
- Hand-edited `account_periods.json` mistakes (non-object entries,
  non-string dates) and malformed sheet dates now raise clean config
  errors or parse as unknown instead of crashing.
- Re-importing an already-imported QIF no longer overcounts `Imported`;
  the dry-run ROC preview uses the same share-balance guard as the write
  path and is labeled "before duplicate filtering".
- The "no transactions could be resolved to an account" fatal now carries
  its `[SKIP]` warnings, and CLI/menu print them before exiting.
- QIF filenames replace Windows-illegal characters instead of crashing
  on accounts whose names contain them.

### Changed
- Personal `ExportToACB.vbs` untracked (repo ships
  `ExportToACB.sample.vbs`); docs sanitized of real account names and
  local paths.

### Added
- `tests/data/sample.qif` example plus round-trip, template-load, and
  duplicate-counter regression tests.

## [1.1.3] — 2026-08-10

### Fixed
- `qif_importer.py`: Duplicate detection used a `set`, which could only
  track whether a (date, type, price, shares) combination existed at all —
  not how many times. Two genuinely distinct transactions sharing the
  exact same date/type/price/shares (e.g. two separate same-size fills at
  the same price on one day) would incorrectly collapse into one, silently
  dropping the second. Switched to a `Counter` (multiset) so each row
  already in the sheet only cancels out one matching QIF transaction.

## [1.1.2] — 2026-08-10

### Fixed
- `qif_importer.py`: Fix dedup key in QIF import to include shares
Previously the duplicate-detection key was (date, type, price),
which caused legitimate same-day/same-price Buy or Sell
transactions with different share counts to be silently skipped
as duplicates. Added shares (rounded to 4dp) to the key so
distinct trades no longer collide.

## [1.1.1] — 2026-05-24

### Fixed
- `ca_calendar.py`: TSX Saturday holidays now correctly observed on the
  preceding Friday instead of the following Monday (US convention was used
  in error). Affected `next_trading_day` / `prev_trading_day` for any year
  where a statutory holiday falls on Saturday (e.g. Christmas 2027,
  New Year's 2028).
- `qif_importer.py`: ROC rows that passed initial validation but later
  failed per-unit price derivation during the write pass were counted in
  both `imported` and `skipped`. Fixed by tracking late-write failures
  separately and deducting them from the `imported` total in the return dict.
- `qimpexp.py`: Switching workbooks (option 7) to a directory without
  `account_periods.json` or `security_map.json` now correctly resets the
  config to empty instead of retaining the previous workbook's mappings.

## [1.1.0] — 2025-04-01
- **Quicken toolbar button integration** (`ExportToACB.vbs`)
  - Single-click automation: drives Quicken's QIF export dialog, then
    calls `qimpexp --import-qif` to update the ACB workbook
  - Pre-flight checks: verifies workbook and qimpexp exist, creates
    output directory, deletes stale QIF to prevent overwrite prompts
  - Export validation: file existence, size, and `!Type:Invst` content
    check before invoking importer
  - Runs `qimpexp` in a visible console window so progress is visible
  - Falls back from `qimpexp.exe` to `python qimpexp.py` automatically
  - Setup instructions in `QUICKEN_SETUP.md`
  
## [1.0.0] — 2025-03-30

### Added
- **Phase 1 — ACB → QIF export**
  - Reads `.xlsx` ACB workbook via `openpyxl`
  - Supports Buy, Sell, ROC (return of capital), and phantom distribution
    transactions
  - Multi-account support via `account_periods.json` — maps each ticker
    to one or more Quicken account names with date ranges; overlap
    validation raises an error on conflicting periods
  - Security name mapping via `security_map.json`
  - Date filters: `--year YYYY`, `--start`, `--end`
  - Mode filters: `full`, `tax-adjustments` (ROC only), `buys-sells`
  - `--dry-run` preview without writing files
  - `--verbose` detailed parse/mapping trace
  - One QIF file per Quicken account, filename
    `<Account>_<YYYYMMDD>_<YYYYMMDD>.qif`
  - QIF date format: Quicken native `M/D'YY` with space-padded day
    (e.g. `8/11'25`, `9/ 2'25`, `12/30'25`)
  - Opening Cash block prepended to each QIF file for correct Quicken
    import starting balance
  - `RtrnCapX` blocks include `L[Account]` and `$amount` lines matching
    Quicken's native export format exactly
  - Amounts formatted with thousands commas (`1,038.48`)

- **ROC share balance logic**
  - Column G (Share Balance) used as authoritative share count for ROC
    amount computation — never column D
  - Column D on ROC rows ignored with a warning if populated
  - Falls back to running balance reconstructed from Buy/Sell rows when
    column G is absent or zero
  - Warning emitted on negative column G; falls back to running balance
  - Warning emitted when share balance goes negative after a Sell
  - Zero-share ROC rows skipped with a warning
  - `_derive_roc_price` finds the shortest decimal precision (4–8 dp)
    that exactly reconstructs the two-decimal-place amount

- **Phase 2 — QIF → ACB import**
  - Reads Quicken-exported QIF files (both native `M/D'YY` and standard
    `MM/DD/YYYY` dates)
  - Supports Buy, Sell, RtrnCapX; other actions skipped with a note
  - Buy/Sell: Quicken trade date → settlement date (T+1 Canadian trading
    day) written to ACB
  - RtrnCapX: record date unchanged (same in both Quicken and ACB)
  - ROC rows with no `I` (price) line: per-unit price derived from
    workbook share balance at the insertion position
  - Rows inserted in correct chronological order, not appended at end
  - Full workbook formula rewrite after insertion preserves ACB
    formula integrity for all rows below the insertion point
  - Duplicate detection with fuzzy price rounding:
    Buy/Sell uses 2 dp (catches Quicken truncation), ROC uses 5 dp
    (distinguishes same-date distributions)
  - Automatic timestamped backup before any write; rollback on formula
    integrity failure
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
