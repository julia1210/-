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
echo  Querying warehouse 20002 inventory...
echo.
".\venv\Scripts\python.exe" query_ecount_inventory.py --wh 20002
echo.
echo  -----------------------------------------------
echo  Press any key to close.
pause >nul
