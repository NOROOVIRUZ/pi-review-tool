@echo off
chcp 65001 >nul
echo ================================================
echo  PI 검토 툴 서버 시작
echo  Made by noroovirus
echo ================================================
echo.

:: 프로젝트 루트로 이동 (이 bat 파일이 deploy\ 안에 있으므로 한 단계 위)
cd /d "%~dp0.."

:: 설치 확인
if not exist ".venv\Scripts\python.exe" (
    echo [오류] 설치가 되어있지 않습니다. deploy\setup_windows.bat 먼저 실행하세요.
    pause
    exit /b 1
)

:: poppler PATH 등록
if exist "poppler\bin" (
    set "PATH=%CD%\poppler\bin;%PATH%"
) else if exist "poppler\Library\bin" (
    set "PATH=%CD%\poppler\Library\bin;%PATH%"
)

:: 포트 설정 (기본 5001)
set PORT=5001
set HOST=0.0.0.0

echo [서버] http://localhost:%PORT% 에서 시작합니다.
echo [서버] 같은 네트워크에서는 http://[이_컴퓨터_IP]:%PORT% 으로 접속 가능
echo.
echo  브라우저를 자동으로 엽니다...
echo  서버를 종료하려면 이 창을 닫으세요.
echo.

:: 3초 후 브라우저 열기 (서버 기동 대기)
start "" cmd /c "timeout /t 3 >nul && start http://localhost:%PORT%"

:: Flask 서버 실행
.venv\Scripts\python app.py
