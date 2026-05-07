@echo off
chcp 65001 >nul
echo ================================================
echo  PI 검토 툴 - Windows 최초 설치
echo  Made by noroovirus
echo ================================================
echo.

:: 프로젝트 루트로 이동
cd /d "%~dp0.."


:: Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python 이 설치되어 있지 않습니다.
    echo https://www.python.org/downloads/ 에서 3.11 이상 설치 후 다시 실행하세요.
    echo 설치 시 "Add Python to PATH" 반드시 체크!
    pause
    exit /b 1
)
echo [OK] Python 확인 완료

:: 가상환경 생성
if not exist ".venv" (
    echo [설치] 가상환경 생성 중...
    python -m venv .venv
)
echo [OK] 가상환경 준비 완료

:: pip 업그레이드 + 패키지 설치
echo [설치] 패키지 설치 중... (처음에는 5~10분 소요)
.venv\Scripts\pip install --upgrade pip >nul 2>&1
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)
echo [OK] 패키지 설치 완료

:: poppler 설치 (pdf2image 의존성)
echo.
echo [설치] poppler 설치 중...
if not exist "poppler" (
    if not exist "poppler-24.08.0.zip" (
        echo     poppler 다운로드 중...
        powershell -Command "Invoke-WebRequest -Uri 'https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip' -OutFile 'poppler-24.08.0.zip'"
    )
    echo     poppler 압축 해제 중...
    powershell -Command "Expand-Archive -Path 'poppler-24.08.0.zip' -DestinationPath 'poppler_temp' -Force"
    :: 최상위 폴더를 poppler 로 이름 변경
    for /d %%i in (poppler_temp\*) do (
        move "%%i" "poppler" >nul 2>&1
    )
    rmdir /s /q poppler_temp 2>nul
    del poppler-24.08.0.zip 2>nul
)
echo [OK] poppler 설치 완료

:: jobs 폴더 생성
if not exist "jobs" mkdir jobs

:: EasyOCR 모델 사전 다운로드 (선택)
echo.
echo [안내] EasyOCR OCR 모델은 첫 번째 스캔 PDF 처리 시 자동 다운로드됩니다.
echo        (~100MB, 첫 실행에만 수 분 소요)
echo.

echo ================================================
echo  설치 완료!
echo  start_windows.bat 을 실행해서 서버를 시작하세요.
echo ================================================
pause
