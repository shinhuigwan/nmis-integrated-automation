@echo off
chcp 65001 >nul
setlocal
set "SCRIPT_DIR=%~dp0"

echo ============================================================
echo   NMIS 통합 자동화 시스템 - 자동 설치
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install.ps1" %*
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (
  echo 설치가 완료되었습니다. 바탕화면의 "통합 자동화 시스템"을 실행하세요.
) else (
  echo 설치 중 오류가 발생했습니다. 위의 오류 내용을 확인하세요.
)
echo.
pause
exit /b %RESULT%
