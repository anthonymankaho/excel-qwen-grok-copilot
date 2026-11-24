@echo off
cd /d "%~dp0"
echo === Qwen for Excel: first-time setup ===
call npm install
call npm run autostart
echo.
echo Close Excel now (if open), then press any key to sideload and start the relay.
pause >nul
call npm start
