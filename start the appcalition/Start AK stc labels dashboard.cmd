@echo off
setlocal

set "APP_DIR=%~dp0.."
set "PORT=4555"
set "NODE_OPTIONS=--use-system-ca"
set "APP_URL=http://localhost:%PORT%/"

cd /d "%APP_DIR%"

netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul
if %ERRORLEVEL% EQU 0 (
  echo AK stc labels dashboard is already running.
  echo Opening %APP_URL%
  start "" "%APP_URL%"
  exit /b 0
)

where node.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo Node.js was not found on this computer.
  echo Install Node.js, then run this file again.
  pause
  exit /b 1
)

echo Starting AK stc labels dashboard...
echo Opening %APP_URL%
echo.
echo Keep this window open while using the dashboard.
start "" /b powershell.exe -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process '%APP_URL%'"
node labels-server.js

echo.
echo The dashboard has stopped.
pause
