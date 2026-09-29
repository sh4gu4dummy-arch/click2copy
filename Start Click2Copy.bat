@echo off
setlocal EnableExtensions
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed or not on PATH.
  echo Install Node.js 18 or later from https://nodejs.org, then try again.
  pause
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo npm is not installed or not on PATH.
  echo Reinstall Node.js 18 or later from https://nodejs.org, then try again.
  pause
  exit /b 1
)

call node "scripts\register-file-association.js"
if errorlevel 1 echo Warning: could not register .c2copy files for double-click opening.

if not exist "node_modules\electron\dist\electron.exe" (
  echo Electron is not installed yet. Setting up the app...
  call npm run setup
  if errorlevel 1 goto :failed
)

echo Starting Click2Copy...
if "%~1"=="" (
  call npm start
) else (
  call npm start -- "%~1"
)
if errorlevel 1 goto :failed
exit /b 0

:failed
echo.
echo Click2Copy could not start. Review the messages above and try again.
pause
exit /b 1
