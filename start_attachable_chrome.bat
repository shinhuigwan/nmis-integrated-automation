@echo off
set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
set "DEBUG_PROFILE=%USERPROFILE%\AppData\Local\nmis-cdp-profile"

if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" exit /b 2

start "" "%CHROME_EXE%" --remote-debugging-port=9222 --user-data-dir="%DEBUG_PROFILE%" "http://nmis.foodservice.or.kr/"
exit /b 0
