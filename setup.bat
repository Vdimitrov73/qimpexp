@echo off
REM ============================================================
REM  QImpExp — Python setup script
REM  Run this once after extracting the source ZIP.
REM  Requires Python 3.9+ with "Add Python to PATH" checked.
REM ============================================================

echo.
echo  QImpExp Setup
echo  -------------

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Python not found on PATH.
    echo  Install Python 3.9+ from https://www.python.org/downloads/
    echo  and make sure "Add Python to PATH" is checked.
    echo.
    pause
    exit /b 1
)

echo  Installing required packages...
python -m pip install --upgrade pip --quiet
python -m pip install openpyxl --quiet

if %errorlevel% neq 0 (
    echo.
    echo  ERROR: pip install failed. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo  Setup complete.
echo.
echo  To run QImpExp:
echo    python qimpexp.py
echo.
echo  Or with options:
echo    python qimpexp.py --acb "C:\path\to\acb_worksheet.xlsx" --year 2025
echo.
pause
