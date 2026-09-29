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

echo Configuring npm to allow Electron's install script...
call npm config set ignore-scripts false
call npm approve-scripts electron 2>nul

echo Installing dependencies (scripts enabled)...
call npm install --foreground-scripts
if errorlevel 1 (
  echo npm install reported an error — continuing to repair Electron...
)

echo Downloading Electron binary directly...
if exist "node_modules\electron\install.js" (
  call node "node_modules\electron\install.js"
) else (
  echo electron package missing — retrying npm install electron...
  call npm install electron --save-dev --foreground-scripts
  if exist "node_modules\electron\install.js" call node "node_modules\electron\install.js"
)

if not exist "node_modules\electron\path.txt" goto :fail_electron
for /f "usebackq delims=" %%A in ("node_modules\electron\path.txt") do set "ELEC_REL=%%A"
if not exist "node_modules\electron\%ELEC_REL%" goto :fail_electron

echo Starting Click2Copy...
call npm start
if errorlevel 1 (
  echo npm start failed.
  pause
  exit /b 1
)
exit /b 0

:fail_electron
echo.
echo Electron binary still missing.
echo Try manually:
echo   npm config set ignore-scripts false
echo   npm approve-scripts electron
echo   rmdir /s /q node_modules
echo   npm install --foreground-scripts
echo   node node_modules\electron\install.js
echo   npm start
echo.
echo If downloads are blocked, check firewall/VPN or set a mirror:
echo   set ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/
pause
exit /b 1
