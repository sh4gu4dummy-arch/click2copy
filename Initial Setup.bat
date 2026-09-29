@echo off
setlocal EnableExtensions
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed or not on PATH.
  echo Install from https://nodejs.org then run this again.
  pause
  exit /b 1
)

echo Installing dependencies...
call npm install --ignore-scripts --no-audit --no-fund > "%TEMP%\click2copy-npm.log" 2>&1
set "NPM_EXIT=%ERRORLEVEL%"
type "%TEMP%\click2copy-npm.log"
if not "%NPM_EXIT%"=="0" goto :fail_install

call :try_electron
if not errorlevel 1 goto :start_app

echo.
echo Still could not download Electron.
echo Logs: %TEMP%\click2copy-npm.log
echo You can open index.html in a browser for UI-only (no save to userData).
pause
exit /b 1

:fail_install
echo Dependency installation failed. See %TEMP%\click2copy-npm.log
pause
exit /b 1

:start_app
echo Starting Click2Copy...
call npm start
if errorlevel 1 (
  echo npm start failed.
  pause
  exit /b 1
)
exit /b 0

:try_electron
call node "scripts\ensure-electron.js" --strict
if errorlevel 1 exit /b 1
if not exist "node_modules\electron\path.txt" exit /b 1
set "ELEC_REL="
for /f "usebackq delims=" %%A in ("node_modules\electron\path.txt") do set "ELEC_REL=%%A"
if "%ELEC_REL%"=="" exit /b 1
if not exist "node_modules\electron\dist\%ELEC_REL%" exit /b 1
echo Electron binary OK: %ELEC_REL%
exit /b 0
