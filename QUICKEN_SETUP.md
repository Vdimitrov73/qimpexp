# Quicken 2009 — Automated ACB Workbook Update

This document explains how to configure Quicken 2009 to automatically
export the current year's QIF and update your ACB workbook with one
button click using `ExportToACB.vbs`.

> The repo ships `ExportToACB.sample.vbs` — copy it to `ExportToACB.vbs`
> (your personal, git-ignored copy) and fill in your values first.
> Below, `ExportToACB.vbs` always means your local copy.

Tested on Quicken Home & Business 2009 R2 Canadian Edition (Windows).
Expected to work on Quicken 2007–2013 for Windows with possible
adjustments to keyboard shortcuts.
Not compatible with Quicken 2014 or later.

---

## How It Works

1. You click the **ExportToACB** toolbar button in Quicken.
2. `ExportToACB.vbs` drives Quicken's built-in QIF export dialog to
   export all investment transactions for the current calendar year
   from **Your Account Name** (use your account) to a fixed file path.
3. The script then runs `qimpexp --import-qif` to compare the exported
   QIF against your ACB workbook and insert any missing Buy, Sell, or
   ROC/Phantom rows — skipping anything already present.

### Date handling

| Transaction | Quicken stores | ACB stores | Conversion |
|---|---|---|---|
| Buy / Sell | Trade date | Settlement date (T+1) | `qimpexp` converts automatically |
| ROC / Phantom | Record date | Record date | No conversion |

---

## Prerequisites

- Quicken Home & Business 2009 (Windows)
- `qimpexp.exe` (or `qimpexp.py` + Python 3.9+) — see [README.md](README.md)
- `security_map.json` placed next to `acb_worksheet.xlsx` — see
  [Security Map](#security-map) below
- `ExportToACB.sample.vbs` (this repo), copied to your personal
  `ExportToACB.vbs` and placed in
  `Documents\QImpExp\`

---

## Step 1 — Place the Script

Copy `ExportToACB.sample.vbs` to `ExportToACB.vbs` at:
C:\Users<YourUsername>\Documents\QImpExp\ExportToACB.vbs

Replace `<YourUsername>` with your actual Windows username.

---

## Step 2 — Configure QHIMENU.INI

Find the file at: C:\ProgramData\Intuit\Quicken\Config\QHIMENU.INI

⚠️ Close Quicken before editing this file.

Add the following section at the **bottom** of the file:

```ini
[QHI]
ExeName=C:\path\to\QImpExp\ExportToACB.vbs
MenuString=ExportToACB
IntuitID=1007
InformExec=FALSE
```

> Replace `C:\path\to\QImpExp` with the folder where you placed the script.

Also find the `AddApps` line (usually near the top of the file) and
add `;QHI` if it is not already present:

```ini
AddApps=;QHI
```

If `AddApps` already has other entries, append `;QHI` to the existing
list, for example:

```ini
AddApps=;QFL;QHI
```

---

## Step 3 — Add the Toolbar Button

1. Start Quicken.
2. Go to **Edit → Customize Toolbar**.
3. Check **Show all choices**.
4. Find **Quicken Home Inventory** (or the item matching IntuitID 1007)
   in the list → click **Add**.
5. Click **Edit Icon** and rename it to `ExportToACB`.

---

## Step 4 — First-Run Checklist

Before clicking the button for the first time:

1. **Verify `security_map.json` exists** next to your workbook — see
   [Security Map](#security-map) below.

2. **Set the Transactions checkbox in Quicken manually once:**
   - Open **File → Export → QIF File** yourself.
   - Tick **Transactions**.
   - Click **Cancel** (not OK — this saves the state without exporting).
   - Quicken remembers this setting for all future automated runs.

3. **Note your account dropdown position:**
   - Open the same dialog and count how many times you press `C` to
     reach **Your Account Name** in the Account dropdown.
   - Set `ACCOUNT_C_PRESSES` in `ExportToACB.vbs` to that number
     (default is `9`).

---

## Step 5 — Daily Usage

Click the **ExportToACB** button in the Quicken toolbar. The script will:

1. Open **File → Export → QIF File** automatically.
2. Set the file path, account, and current-year date range.
3. Export the QIF to `Documents\Tax Documents\QImpExp\Your Account Name.QIF`.
4. Validate the exported file (size and format checks).
5. Run `qimpexp --import-qif` in a console window.
6. Show per-transaction detail (`[WRITE]` / `[SKIP-DUP]` / `[WARN]`).
7. Display a completion message box when done.

> The full process takes approximately 15–30 seconds depending on
> workbook size and machine speed.

---

## Security Map

`security_map.json` maps your workbook tickers to the exact security
names used in Quicken. Place it next to your workbook:
C:\path\to\Tax Documents\security_map.json

text

Example:

```json
{
  "VBAL": "VANGUARD BAL ETF PTFL",
  "ZCN":  "BMO S&P/TSX CAPP COMP ETF",
  "CPD":  "ISHRS S&P/TSX CDN PFD ETF",
  "VRE":  "VANGUARD FTSE CDN CAP REIT ETF",
  "MNT":  "ROYAL CDN MINT - CDN GOLD"
}
```

Security names must match Quicken's Security List exactly (case-sensitive).

---

## What ExportToACB Does NOT Import

Quicken exports several action types that have no ACB impact and are
intentionally skipped:

| Quicken action | What it is | ACB action |
|---|---|---|
| `Div` | Gross distribution paid out | Skipped — ROC breakdown comes from T3/T5 slips |
| `XOut` | Internal cash transfer between accounts | Skipped — no shares, no ACB effect |
| `RtrnCapX` | Return of capital (if present) | Imported as ROC row |

ROC and phantom distribution rows in your ACB workbook come from your
annual T3/T5 tax slips, not from Quicken's `Div` entries. After you
receive your T-slips each year, use the qimpexp interactive menu
(option 3 — Import QIF, mode `tax-adjustments`) to add those rows.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Wrong account selected | `ACCOUNT_C_PRESSES` is wrong | Count `C` presses manually; update the constant |
| QIF exported with 0 bytes | Transactions checkbox not ticked | Run the export dialog manually once and tick it |
| "security_map.json required" | File missing or wrong location | Place it next to `acb_worksheet.xlsx` |
| "QIF export failed — file not created" | Menu accelerator wrong | Check `MENU_EXPORT` / `MENU_QIF` constants |
| All rows show `[SKIP-DUP]` | Transactions already in workbook | Normal — the workbook is up to date |
| Script clicks wrong thing | Quicken was slow to respond | Increase the relevant `WScript.Sleep` value |
| SmartScreen warning on `qimpexp.exe` | Unsigned executable | Click **More info → Run anyway**; see BUILD_EXE.md |

---

## Files in This Repo

| File | Description |
|---|---|
| `qimpexp.exe` | Standalone Windows executable (also on GitLab/GitHub Releases) |
| `qimpexp.py` | CLI entry point and interactive menu |
| `ExportToACB.sample.vbs` | Sample script — copy to personal `ExportToACB.vbs`; automates Quicken export and triggers qimpexp import |
| `security_map.json` | Ticker → Quicken security name mapping (place next to workbook) |
| `account_periods.json` | Ticker → Quicken account name mapping with date ranges |
| `README.md` | Full project documentation |
| `QUICKEN_SETUP.md` | This file |
