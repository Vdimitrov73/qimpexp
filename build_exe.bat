@echo off
REM build_exe.bat - Local production build of qimpexp.exe from the working tree.
REM Mirrors the flags of the build-exe job in .gitlab-ci.yml.
REM Run from the repo folder:  build_exe.bat
REM The stale qimpexp.exe (if any) is backed up first; outputs are git-ignored.
setlocal EnableDelayedExpansion
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (echo ERROR: python not found on PATH. & exit /b 1)

REM Always build in a clean venv (like CI) so site-packages from the
REM current environment cannot bloat the exe.
if not exist "build_env\Scripts\python.exe" (
  echo Creating clean build venv...
  python -m venv build_env || exit /b 1
)
echo Installing build deps into venv...
build_env\Scripts\python -m pip install openpyxl pyinstaller || exit /b 1
set PY=build_env\Scripts\python

if exist "qimpexp.exe" (
  for /f %%t in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%t
  echo Backing up stale exe to qimpexp_backup_!TS!.exe
  copy /y "qimpexp.exe" "qimpexp_backup_!TS!.exe" >nul || exit /b 1
)

echo Building qimpexp.exe from working tree...
%PY% -m PyInstaller --onefile --console --name qimpexp ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import openpyxl.workbook ^
  --version-file version.txt ^
  qimpexp.py || exit /b 1

copy /y "dist\qimpexp.exe" "qimpexp.exe" >nul || exit /b 1
echo Replaced: qimpexp.exe

qimpexp.exe --help >nul || (
  echo New exe still locked (antivirus scan?) - waiting 15s and retrying...
  timeout /t 15 /nobreak >nul
  qimpexp.exe --help >nul || (echo ERROR: fresh exe failed to launch. & exit /b 1)
)
echo Smoke test OK (qimpexp.exe --help launches).
