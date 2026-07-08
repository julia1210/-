@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    ECOUNT Kit - First Install
echo ============================================
echo.
echo [1/3] Checking Python...
python --version >nul 2>&1
if errorlevel 1 goto NOPYTHON
echo       OK
echo.
echo [2/3] Creating venv... (1-2 min)
python -m venv venv
if errorlevel 1 goto VENVFAIL
echo       OK
echo.
echo [3/3] Installing libraries...
".\venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".\venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto PIPFAIL
echo.
echo ============================================
echo    Done! Run 2_test.bat next.
echo ============================================
pause
exit /b 0

:NOPYTHON
echo.
echo [ERROR] Python is not installed.
echo   1) Go to https://www.python.org/downloads/
echo   2) Download and install Python
echo   3) CHECK "Add Python to PATH" during install
echo   4) Run this file again
echo.
pause
exit /b 1

:VENVFAIL
echo.
echo [ERROR] Failed to create venv.
pause
exit /b 1

:PIPFAIL
echo.
echo [ERROR] Failed to install libraries. Check internet connection.
pause
exit /b 1
