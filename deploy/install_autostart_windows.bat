@echo off
chcp 65001 >nul
echo ================================================
echo  PI 검토 툴 - Windows 시작 시 자동 실행 등록
echo ================================================
echo.
echo 이 스크립트를 실행하면 Windows 시작 시 PI 검토 툴이
echo 자동으로 백그라운드에서 실행됩니다.
echo.

:: 현재 폴더 경로 (bat 파일 위치 기준 상위 폴더 = 프로젝트 루트)
set "PROJECT_DIR=%~dp0.."
set "START_BAT=%~dp0start_windows.bat"

:: VBScript로 숨김 창 실행 래퍼 생성 (콘솔 창 숨기기)
set "VBS_PATH=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\pi_review_tool.vbs"

echo Set oShell = CreateObject("WScript.Shell") > "%VBS_PATH%"
echo oShell.CurrentDirectory = "%PROJECT_DIR%" >> "%VBS_PATH%"
echo oShell.Run """%START_BAT%""", 0, False >> "%VBS_PATH%"

echo [OK] 시작 프로그램에 등록했습니다.
echo      위치: %VBS_PATH%
echo.
echo 제거하려면 해당 파일을 삭제하세요:
echo %VBS_PATH%
echo.
pause
