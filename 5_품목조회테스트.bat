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
echo  Testing item list API...
echo  (Check output/item_api_raw.json for full response)
echo.
".\venv\Scripts\python.exe" test_item_api.py
echo.
echo  -----------------------------------------------
echo  Press any key to close.
pause >nul
