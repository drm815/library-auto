@echo off
chcp 65001 > nul
echo.
echo ========================================
echo  학교도서관 자동화 시스템 설치 프로그램
echo ========================================
echo.

:: Python 설치 확인
echo [1/6] Python 설치 확인 중...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ⚠️  Python이 설치되어 있지 않습니다!
    echo.
    if exist "%~dp0python-installer.exe" (
        echo 📦 Python 설치 파일을 발견했습니다. 자동으로 설치합니다...
        echo.
        echo ※ 설치 창이 열리면 아래를 반드시 체크하세요:
        echo    [Add Python to PATH] 체크 ✅
        echo    그 다음 [Install Now] 클릭!
        echo.
        pause
        "%~dp0python-installer.exe"
        echo.
        echo Python 설치 완료 후 아무 키나 누르세요...
        pause
    ) else (
        echo.
        echo ❌ Python 설치 파일을 찾을 수 없습니다!
        echo 같은 폴더에 python-installer.exe 파일이 있는지 확인해주세요.
        pause
        exit
    )
    python --version > nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo ❌ Python 설치가 완료되지 않았습니다.
        echo 설치 시 [Add Python to PATH] 를 체크했는지 확인하고 다시 시도해주세요.
        pause
        exit
    )
)
for /f "tokens=*" %%i in ('python --version') do echo ✅ %%i 확인됨

echo.
echo [2/6] 필요한 라이브러리 설치 중... (시간이 걸릴 수 있어요)
echo.
pip install selenium > nul 2>&1
echo ✅ selenium 설치 완료
pip install webdriver-manager > nul 2>&1
echo ✅ webdriver-manager 설치 완료
pip install google-auth google-auth-httplib2 google-api-python-client > nul 2>&1
echo ✅ google-auth 설치 완료
pip install google-auth-oauthlib > nul 2>&1
echo ✅ google-auth-oauthlib 설치 완료
pip install gspread > nul 2>&1
echo ✅ gspread 설치 완료
pip install pandas openpyxl > nul 2>&1
echo ✅ pandas, openpyxl 설치 완료

echo.
echo [3/6] Python 경로 확인 중...
for /f "tokens=*" %%i in ('where python') do (
    set PYTHON_PATH=%%i
    goto :found
)
:found
echo ✅ Python 경로: %PYTHON_PATH%

echo.
echo [4/6] 파일 위치 확인 중...
set SCRIPT_PATH=%~dp0library_auto.py
if not exist "%SCRIPT_PATH%" (
    echo ❌ library_auto.py 파일을 찾을 수 없습니다!
    echo 같은 폴더에 library_auto.py 파일이 있는지 확인해주세요.
    pause
    exit
)
echo ✅ library_auto.py 확인됨: %SCRIPT_PATH%

echo.
echo [5/6] 백그라운드 실행 파일 생성 중...
set VBS_PATH=%~dp0run_hidden.vbs
echo Set WshShell = CreateObject("WScript.Shell") > "%VBS_PATH%"
echo WshShell.Run """%PYTHON_PATH%"" ""%SCRIPT_PATH%""", 0, False >> "%VBS_PATH%"
echo ✅ run_hidden.vbs 생성 완료

echo.
echo [6/6] 작업 스케줄러 등록 중...
schtasks /delete /tn "도서관자동화" /f > nul 2>&1
schtasks /create /tn "도서관자동화" /tr "wscript \"%VBS_PATH%\"" /sc hourly /mo 1 /st 08:00 /f > nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ 작업 스케줄러 등록 완료! ^(매일 오전 8시부터 1시간마다 자동 실행^)
) else (
    echo ⚠️  작업 스케줄러 등록 실패 - 관리자 권한으로 다시 실행해보세요.
)

echo.
echo ========================================
echo  설치 완료!
echo ========================================
echo.
echo 다음 단계:
echo   1. library_auto.py 파일에 독서로 계정 정보 입력
echo   2. python library_auto.py 실행하여 구글 인증
echo   3. 이후부터는 1시간마다 백그라운드에서 자동 실행!
echo.
pause
