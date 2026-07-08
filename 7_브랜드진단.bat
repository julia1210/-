@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
if not exist ".\venv\Scripts\python.exe" (
  echo Run 1_install.bat first.
  pause
  exit /b 1
)
echo.
echo  Checking brand (CONT1) values in item master...
echo.
".\venv\Scripts\python.exe" debug_brands.py
echo.
echo  -----------------------------------------------
echo  Press any key to close.
pause >nul
