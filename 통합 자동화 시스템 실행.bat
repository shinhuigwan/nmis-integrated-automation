@echo off
chcp 65001 >nul
if "%1"=="hide" goto :run
mshta vbscript:createobject("wscript.shell").run("""%~f0"" hide",0)(window.close)&exit /b

:run
setlocal
set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"
set "VENV_PYTHONW=%VENV_DIR%\Scripts\pythonw.exe"

where python >nul 2>&1
if errorlevel 1 exit /b 2

if not exist "%VENV_PYTHON%" (
  python -m venv "%VENV_DIR%"
  if errorlevel 1 goto :install_error
  "%VENV_PYTHON%" -m pip install -r "%SCRIPT_DIR%requirements.txt"
  if errorlevel 1 goto :install_error
)

if exist "%VENV_PYTHONW%" (
  "%VENV_PYTHONW%" "%SCRIPT_DIR%nmis_slip_ui.py" %*
) else (
  "%VENV_PYTHON%" "%SCRIPT_DIR%nmis_slip_ui.py" %*
)
exit /b 0

:install_error
exit /b 1
