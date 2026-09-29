@echo off
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed or not on PATH.
  echo Install from https://nodejs.org then run this again.
  pause
  exit /b 1
)
echo Installing dependencies...
call npm install
if errorlevel 1 (
  echo npm install failed.
  pause
  exit /b 1
)
echo Starting Click2Copy...
call npm start
if errorlevel 1 (
  echo npm start failed.
  pause
  exit /b 1
)
