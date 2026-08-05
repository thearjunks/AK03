@echo off
setlocal

set "APP_DIR=C:\Users\thear\OneDrive\Documents\New project"
set "PORT=4555"
set "NODE_OPTIONS=--use-system-ca"
set "APP_URL=http://localhost:%PORT%/"

cd /d "%APP_DIR%"

netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul
if %ERRORLEVEL% EQU 0 (
  echo AK stc labels dashboard is already running.
  echo Opening %APP_URL%
  start "" "%APP_URL%"
  pause
  exit /b 0
)

echo Starting AK stc labels dashboard on %APP_URL%
echo Keep this window open while using the dashboard.
start "" cmd /c "timeout /t 2 >nul & start "" "%APP_URL%""
node labels-server.js

echo.
echo Dashboard stopped.
pause
