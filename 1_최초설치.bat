@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    이카운트 연결 키트 - 최초 설치 (처음 1회)
echo ============================================
echo.
echo [1/3] 파이썬(Python) 설치 확인...
python --version >nul 2>&1
if errorlevel 1 goto NOPYTHON
echo       OK
echo.
echo [2/3] 가상환경(venv) 생성... (1~2분 걸립니다)
python -m venv venv
if errorlevel 1 goto VENVFAIL
echo       OK
echo.
echo [3/3] 필요한 라이브러리 설치...
".\venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".\venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto PIPFAIL
echo.
echo ============================================
echo    설치 완료!  이제 '2_연결테스트.bat' 를 실행하세요.
echo ============================================
pause
exit /b 0

:NOPYTHON
echo.
echo [중요] 이 PC에 파이썬(Python)이 설치되어 있지 않습니다.
echo.
echo   1) https://www.python.org/downloads/  접속
echo   2) 노란색 Download 버튼으로 설치 파일을 받습니다
echo   3) 설치 첫 화면에서 "Add Python to PATH" 를 꼭 체크하세요 (★중요★)
echo   4) 설치가 끝나면 이 파일(1_최초설치.bat)을 다시 실행하세요
echo.
pause
exit /b 1

:VENVFAIL
echo.
echo [오류] 가상환경 생성에 실패했습니다. 본사 담당자에게 이 화면을 캡처해 문의하세요.
pause
exit /b 1

:PIPFAIL
echo.
echo [오류] 라이브러리 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 실행하세요.
pause
exit /b 1
