@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js 18 or newer is required. Install it from https://nodejs.org/
  pause
  exit /b 1
)

start "" /min powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:4173'"

echo Starting local NVIDIA chat on http://127.0.0.1:4173
echo Press Ctrl+C to stop.
echo.
node server.js

echo.
echo Server stopped.
pause
