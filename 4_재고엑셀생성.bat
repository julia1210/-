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
echo  Building inventory Excel... (may take 1-2 min)
echo.
".\venv\Scripts\python.exe" build_inventory_excel.py
echo.
echo  -----------------------------------------------
echo  Press any key to close.
pause >nul
