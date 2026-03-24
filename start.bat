@echo off
title Ctrip Visa Scraper

echo.
echo  ==========================================
echo   Ctrip Visa Order Scraper - Starting...
echo  ==========================================
echo.

cd /d "%~dp0"

echo [Step 1/2] Installing dependencies...
echo.
call npm install
if errorlevel 1 (
    echo.
    echo  ERROR: npm install failed!
    pause
    exit /b 1
)

echo.
echo [Step 2/2] Starting server...
echo.
echo  ----------------------------------------
echo  Open browser: http://localhost:3333
echo  Do NOT close this window!
echo  ----------------------------------------
echo.

node server.js

pause
