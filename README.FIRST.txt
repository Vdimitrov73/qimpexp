============================================================
  QImpExp — Quick Start
============================================================

STEP 1 — Run the tool
-----------------------
Double-click qimpexp.exe.

On first launch, choose whether to run in interactive
menu mode (no arguments) or CLI mode.

For interactive mode, just double-click the exe and
follow the prompts.


STEP 2 — First-time setup wizard
----------------------------------
The first time you point QImpExp at a workbook folder
that has no config files, it runs a setup wizard asking:

  1. Security names
     The exact Quicken security name for each fund ticker
     (e.g. "VANGUARD BAL ETF PTFL" for VBAL).
     Press Enter to use the ticker symbol as-is.

  2. Account mapping
     The Quicken account name where each fund is held
     (e.g. "Your Account Name").
     If a fund moved between accounts during the year,
     you can add multiple date periods.

After the wizard, two config files are saved next to
your ACB workbook:

  account_periods.json    — ticker-to-account mapping
  security_map.json       — ticker-to-security-name mapping

These files are excluded from git via .gitignore.
You can edit them manually at any time.


STEP 3 — Export ACB workbook to QIF (Phase 1)
-----------------------------------------------
From the interactive menu, choose option 1 (Export).
Or from the command line:

  qimpexp.exe --acb "C:\path\to\acb_worksheet.xlsx" --year 2025

One QIF file is written per Quicken account, named:
  <Account>_<YYYYMMDD>_<YYYYMMDD>.qif

Preview without writing (dry run):
  qimpexp.exe --acb acb_worksheet.xlsx --year 2025 --dry-run

Common options:
  --funds VBAL ZCN        Process only specific tickers
  --mode tax-adjustments  ROC transactions only
  --mode buys-sells       Buy/Sell only
  --output C:\QIF\        Write files to a specific folder
  --verbose               Show detailed parsing decisions


STEP 4 — Import QIF into ACB workbook (Phase 2)
-------------------------------------------------
To import a Quicken-exported QIF file back into your
ACB workbook, choose option 3 (Import) from the menu.
Or from the command line:

  qimpexp.exe --acb acb_worksheet.xlsx --import-qif Export.QIF

  *** Always preview with --dry-run before the first import ***

  qimpexp.exe --acb acb_worksheet.xlsx --import-qif Export.QIF --dry-run

Buy/Sell trade dates are automatically converted to
settlement dates (T+1 Canadian trading day).
ROC record dates are imported unchanged.
Reinvested dividends (a Quicken Buy whose memo
begins with DRIP, e.g. Drip VRE) are imported
without the T+1 shift and marked DRIP in
column K of the sheet.

After import, open the workbook and verify that the
formulas in columns F-J look correct before saving.


STEP 5 — Import QIF into Quicken
----------------------------------
In Quicken: File > Import > QIF File...
Select the account, then browse to the .qif file.

If Quicken shows "Security not found", make sure the
security name in security_map.json exactly matches
the name shown in Quicken's Security List.


============================================================
  Command-line reference
============================================================

  qimpexp.exe --acb PATH [options]

  --acb PATH            Path to ACB .xlsx workbook (required)
  --year YYYY           Filter to a full calendar year
  --start YYYY-MM-DD    Start date filter (inclusive)
  --end   YYYY-MM-DD    End date filter (inclusive)
  --funds FUND ...      Limit to specific tickers
  --mode  full|tax-adjustments|buys-sells
  --output PATH         Output directory (default: current dir)
  --dry-run             Preview without writing files
  --verbose             Detailed trace output
  --import-qif PATH     Import a QIF file into the ACB workbook


============================================================
  Need help?
============================================================

Full documentation:
  https://github.com/Vdimitrov73/qimpexp
  https://gitlab.com/vdimitrov_73/qimpexp

Report an issue:
  https://github.com/Vdimitrov73/qimpexp/issues
  https://gitlab.com/vdimitrov_73/qimpexp/-/issues

SmartScreen warning ("Windows protected your PC"):
  Click "More info" then "Run anyway". This appears
  because the exe is not signed with a paid EV cert
  (~$500 USD/year). The tool is open source — you
  can inspect every line of code on GitHub/GitLab,
  or build the exe yourself (see BUILD_EXE.md).


============================================================
  Disclaimer
============================================================

This tool is for personal convenience only. It is not
tax or financial advice. Always cross-check generated
QIF files against your broker statements before
importing into Quicken. Consult a tax professional
if in doubt.

============================================================
