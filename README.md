# QImpExp — ACB Workbook ↔ Quicken QIF Converter

QImpExp converts your ACB (Adjusted Cost Base) Excel workbook into
Quicken-compatible QIF files for investment accounts — and imports
Quicken-exported QIF files back into your ACB workbook.

Handles Buy, Sell, ROC (Return of Capital), and phantom distribution
transactions, with full support for multi-account, multi-fund setups
and Canadian TSX settlement dates.

---

## Requirements

- Python 3.9 or later
- [`openpyxl`](https://pypi.org/project/openpyxl/) — the only third-party dependency

```
pip install openpyxl
```

Or use the standalone Windows `.exe` (no Python required) — see [Installation](#installation).

---

## Installation

### Option A — Standalone Windows exe (no Python required)

Download the latest release ZIP from
[GitHub Releases](https://github.com/Vdimitrov73/qimpexp/releases/latest)
or [GitLab Releases](https://gitlab.com/vdimitrov_73/qimpexp/-/releases/permalink/latest),
unzip anywhere, and run `qimpexp.exe`.

> **Windows SmartScreen warning?** Click **"More info"** → **"Run anyway"**.
> See [BUILD_EXE.md](BUILD_EXE.md) for why this happens and how to build from source.

### Option B — Python (developers / all platforms)

```
git clone https://github.com/Vdimitrov73/qimpexp.git
cd qimpexp
pip install openpyxl
python qimpexp.py
```

---

## Quick Start

Run with no arguments for the interactive menu:

```
python qimpexp.py
```

Or use the CLI directly:

```
python qimpexp.py --acb acb_worksheet.xlsx --year 2025
```

On first run, a setup wizard collects your security names and account
mapping, then saves them as `account_periods.json` and `security_map.json`
next to your workbook.

---

## Quicken Toolbar Button Integration

QImpExp includes `ExportToACB.vbs`, a VBScript that automates the full
export → import cycle from a single Quicken toolbar button:

1. Drives Quicken's **File → Export → QIF File** dialog automatically.
2. Exports the current year's transactions for your investment account.
3. Runs `qimpexp --import-qif` to insert any missing rows into the ACB
   workbook — skipping duplicates safely.

See **[QUICKEN_SETUP.md](QUICKEN_SETUP.md)** for setup instructions.

### `ExportToACB.vbs` configuration

Open the script in any text editor and adjust these constants at the top:

| Constant | Default | What to change |
|---|---|---|
| `ACCOUNT_C_PRESSES` | `9` | Number of times to press `C` in the account dropdown to reach your account |
| `MENU_EXPORT` | `"e"` | Accelerator key for **Export** in the File menu |
| `MENU_QIF` | `"q"` | Accelerator key for **QIF File...** in the Export submenu |
| `baseDir` | `%USERPROFILE%\Documents\Tax Documents\QImpExp\` | Folder containing the QIF export and qimpexp |
| `ACB_PATH` | `%USERPROFILE%\Documents\Tax Documents\acb_worksheet.xlsx` | Path to your ACB workbook |

---

## Project Structure

```
qimpexp.py              CLI entry point and interactive menu
parsers.py              Workbook reading, config loading, account/security resolution
qif_writer.py           QIF building and file writing
qif_importer.py         Phase 2 — QIF → ACB reverse import
ca_calendar.py          Canadian TSX market calendar (T+1 settlement)
qif_colors.py           ANSI colour utilities for the interactive menu
ExportToACB.vbs         Quicken toolbar button script (export + import automation)
version.txt             PyInstaller version resource (maintainer use)
QUICKEN_SETUP.md        Quicken toolbar button setup guide
account_periods.json  Example account period config
security_map.json     Example security name mapping
```

---

## Phase 1 — Export ACB Workbook to QIF

### CLI usage

```
python qimpexp.py --acb PATH [options]
```

| Argument | Description |
|----------|-------------|
| `--acb PATH` | Path to ACB `.xlsx` workbook (required) |
| `--year YYYY` | Shortcut for `--start YYYY-01-01 --end YYYY-12-31` |
| `--start YYYY-MM-DD` | Start date filter (inclusive) |
| `--end YYYY-MM-DD` | End date filter (inclusive) |
| `--funds FUND ...` | Limit to specific tickers (e.g. `VBAL ZCN`) |
| `--mode` | `full` (default) · `tax-adjustments` · `buys-sells` |
| `--output PATH` | Output directory (default: current directory) |
| `--dry-run` | Print QIF to stdout without writing files |
| `--verbose` | Show detailed parsing and mapping decisions |
| `--account NAME` | Default account name (fallback when `account_periods.json` is absent) |

### Examples

```bash
# All transactions for 2025
python qimpexp.py --acb acb_worksheet.xlsx --year 2025

# Only VBAL and ZCN, write to a specific folder
python qimpexp.py --acb acb_worksheet.xlsx --year 2025 --funds VBAL ZCN --output C:\QIF\

# ROC transactions only (for tax-year ACB adjustment import)
python qimpexp.py --acb acb_worksheet.xlsx --year 2025 --mode tax-adjustments

# Preview without writing
python qimpexp.py --acb acb_worksheet.xlsx --year 2025 --dry-run --verbose
```

---

## Phase 2 — Import QIF into ACB Workbook

```bash
# Preview what would be appended (recommended first step)
python qimpexp.py --acb acb_worksheet.xlsx --import-qif Export.QIF --dry-run

# Import
python qimpexp.py --acb acb_worksheet.xlsx --import-qif Export.QIF
```

### Date conversion

Canadian securities settle T+1. QImpExp handles the date translation
automatically:

| Transaction type | QIF date | ACB date |
|-----------------|----------|----------|
| Buy / Sell | Trade date | Settlement date (next TSX trading day) |
| RtrnCapX (ROC) | Record date | Record date (unchanged) |

The Canadian market calendar (`ca_calendar.py`) accounts for all TSX
statutory holidays including Good Friday, Victoria Day, and the Christmas
break.

---

## Input Workbook Format

Each worksheet in the ACB workbook is named after a fund ticker
(e.g. `VBAL`, `ZCN`, `CPD`, `VRE`). Sheet names are normalized to
uppercase, so `CASH-Vlad` is treated as ticker `CASH-VLAD`.

| Column | Field | Notes |
|--------|-------|-------|
| A | Date | Any parseable date format; also handles text like `2025-Apr-01` |
| B | Type | `Buy`, `Sell`, or `ROC` (case-insensitive) |
| C | Price per share | For ROC rows: ROC per-unit amount |
| D | Shares | For ROC rows: **ignored** — column G is used instead |
| E | Commission | Optional |
| F | Capital Gains | Calculated — ignored on read |
| **G** | **Share Balance** | **Used as authoritative share count for ROC** |
| H–J | ACB Change / ACB / ACB/Share | Calculated — ignored on read |

> **Important:** For ROC rows, QImpExp always uses column G (Share Balance)
> as the share count. Column D is ignored even if populated — a warning is
> emitted if column D has a value on a ROC row. This prevents incorrect ROC
> amounts when the Shares column was accidentally filled in.

---

## Configuration Files

Place both files in the same directory as your workbook.
The interactive setup wizard creates them for you on first run.

### `account_periods.json`

Maps each ticker to one or more Quicken account names with date ranges.
Useful when a fund moved between accounts during its history.

```json
{
  "VBAL": [
    {"account": "Your Account A", "start": "2021-01-01", "end": "2025-08-10"},
    {"account": "Your Account Name",      "start": "2025-08-11"}
  ],
  "ZCN": [
    {"account": "Your Account A", "start": "2021-01-01", "end": "2025-08-10"},
    {"account": "Your Account Name",      "start": "2025-08-11"}
  ]
}
```

- `end` is optional — omit it for an open-ended (current) period.
- If this file is missing, QImpExp falls back to `--account` or prompts interactively.
- QImpExp never guesses or invents account names.

### `security_map.json`

Maps tickers to the exact Quicken security names shown in your Security List.

```json
{
  "VBAL": "VANGUARD BAL ETF PTFL",
  "ZCN":  "BMO S&P/TSX CAPP COMP ETF",
  "CPD":  "ISHRS S&P/TSX CDN PFD ETF",
  "VRE":  "VANGUARD FTSE CDN CAP REIT ETF",
  "MNT":  "ROYAL CDN MINT - CDN GOLD"
}
```

- If missing, QImpExp prompts interactively. Press Enter to use the ticker as-is.
- Security names must match Quicken exactly for the import to find the right security.

---

## QIF Output Format

One QIF file is written per Quicken account.

**Filename:** `<Account>_<YYYYMMDD>_<YYYYMMDD>.qif`

```
!Type:Invst
L[Your Account Name]

D8/11'25
NBuy
YVANGUARD BAL ETF PTFL
I31.877597
Q12437
U396,461.67
T396,461.67
^

D9/ 2'25
NRtrnCapX
YBMO S&P/TSX CAPP COMP ETF
U26.83
T26.83
L[Your Account Name]
$26.83
^
```

### Action mapping

| Workbook type | QIF action | T/U amount |
|---------------|------------|------------|
| Buy | `Buy` | price × shares + commission |
| Sell | `Sell` | price × shares − commission |
| ROC (per-unit > 0) | `RtrnCapX` | per-unit × share balance |
| ROC (per-unit < 0) | `RtrnCapX` | negative amount (phantom distribution) |

`RtrnCapX` blocks include `L[Account]` and `$amount` lines, matching
Quicken's native export format exactly.

---

## Output Modes

| Mode | Includes |
|------|----------|
| `full` | Buy, Sell, ROC (default) |
| `tax-adjustments` | ROC only |
| `buys-sells` | Buy and Sell only |

---

## Warnings and Errors

**Fatal errors** (exit immediately):
- Workbook cannot be opened
- No valid sheets found after fund filter
- No transactions remain after date filtering
- Malformed `account_periods.json`

**Warnings** (logged; processing continues):
- Unparseable date in a row
- Unsupported transaction type
- Missing price or shares for Buy/Sell
- ROC row has a value in column D (col D ignored; col G used instead)
- Share balance went negative after a Sell
- Zero-share ROC row (skipped)
- Transaction has no matching account period (skipped)

---

## Building the .exe

See [BUILD_EXE.md](BUILD_EXE.md) for instructions on building
`qimpexp.exe` with PyInstaller, plus an explanation of the Windows
SmartScreen warning.

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

---

## License

MIT License. See [LICENSE](LICENSE) for the full text.

---

## Disclaimer

This tool is for personal convenience only. It is not tax or financial
advice. Always cross-check generated QIF files against your broker
statements before importing into Quicken.
