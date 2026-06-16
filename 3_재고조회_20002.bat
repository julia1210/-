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
echo  이카운트 20002(디아이 판매) 창고 재고현황을 조회합니다...
echo.
".\venv\Scripts\python.exe" query_ecount_inventory.py --wh 20002
echo.
echo  -----------------------------------------------
echo  창을 닫으려면 아무 키나 누르세요.
pause >nul
