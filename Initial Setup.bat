@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed or not on PATH.
  echo Install from https://nodejs.org then run this again.
  pause
  exit /b 1
)

echo Approving Electron install scripts (needed on newer npm)...
call npm approve-scripts electron 2>nul
call npm config set ignore-scripts false

echo Installing dependencies...
call npm install
if errorlevel 1 (
  echo npm install failed.
  pause
  exit /b 1
)

echo Ensuring Electron binary is present...
call node scripts\ensure-electron.js
if errorlevel 1 (
  echo.
  echo Trying a clean Electron reinstall...
  if exist "node_modules\electron" rmdir /s /q "node_modules\electron"
  call npm approve-scripts electron 2>nul
  call npm install electron --save-dev
  call node scripts\ensure-electron.js
  if errorlevel 1 (
    echo.
    echo Still failing. Manual steps:
    echo   1. npm approve-scripts electron
    echo   2. rmdir /s /q node_modules\electron
    echo   3. npm install electron --save-dev
    echo   4. npm start
    pause
    exit /b 1
  )
)

echo Starting Click2Copy...
call npm start
if errorlevel 1 (
  echo npm start failed.
  pause
  exit /b 1
)
