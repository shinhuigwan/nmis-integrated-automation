@echo off
setlocal
set "SCRIPT_DIR=%~dp0"

echo ============================================================
echo   NMIS Integrated Automation - One-click setup
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install.ps1" %*
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (
  echo Setup completed. Use the desktop shortcut to start the app.
) else (
  echo Setup failed. Review the error message above.
)
echo.
pause
exit /b %RESULT%
