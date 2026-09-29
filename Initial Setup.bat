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

call npm config set ignore-scripts false
call npm approve-scripts electron 2>nul

REM Prefer official; if binary still missing, retry with public mirror
set "ELECTRON_GET_USE_PROXY=0"

echo Installing dependencies...
call npm install --foreground-scripts > "%TEMP%\click2copy-npm.log" 2>&1
type "%TEMP%\click2copy-npm.log"

call :try_electron
if not errorlevel 1 goto :start_app

echo Official Electron download failed — retrying with mirror...
set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
if exist "node_modules\electron" rmdir /s /q "node_modules\electron"
call npm install electron --save-dev --foreground-scripts > "%TEMP%\click2copy-electron.log" 2>&1
type "%TEMP%\click2copy-electron.log"
call :try_electron
if not errorlevel 1 goto :start_app

echo.
echo Still could not download Electron.
echo Logs: %TEMP%\click2copy-npm.log and %TEMP%\click2copy-electron.log
echo You can open index.html in a browser for UI-only (no save to userData).
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
if not exist "node_modules\electron\install.js" exit /b 1
echo Running electron install.js ...
call node "node_modules\electron\install.js" > "%TEMP%\click2copy-elec-install.log" 2>&1
type "%TEMP%\click2copy-elec-install.log"
if not exist "node_modules\electron\path.txt" exit /b 1
set "ELEC_REL="
for /f "usebackq delims=" %%A in ("node_modules\electron\path.txt") do set "ELEC_REL=%%A"
if "%ELEC_REL%"=="" exit /b 1
if not exist "node_modules\electron\%ELEC_REL%" exit /b 1
echo Electron binary OK: %ELEC_REL%
exit /b 0
