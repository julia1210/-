@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
if not exist ".\venv\Scripts\python.exe" (
  echo 먼저 '1_최초설치.bat' 를 실행해 주세요.
  pause
  exit /b 1
)
echo.
echo  이카운트 전 창고 재고를 조회해 Excel 파일을 생성합니다...
echo  (창고 수가 많아 1~2분 걸릴 수 있습니다)
echo.
".\venv\Scripts\python.exe" build_inventory_excel.py
echo.
echo  -----------------------------------------------
echo  창을 닫으려면 아무 키나 누르세요.
pause >nul
