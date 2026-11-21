@echo off
title The Indian Express - Public Live URL Server
color 0B
cls
echo ======================================================================
echo       THE INDIAN EXPRESS - PUBLIC LIVE URL GENERATOR
echo ======================================================================
echo.
echo  Starting local server and generating a Public HTTPS URL...
echo  (Is URL ko aap kisi bhi team member ko bhej sakte hain!)
echo.
echo ======================================================================

cd /d "%~dp0"

:: Start local web server in background
start /b python web_app.py

:: Wait 2 seconds for server to start
timeout /t 2 /nobreak >nul

echo.
echo Local server running on http://localhost:5000
echo.
echo Connecting to secure tunnel for Public URL...
echo (Public HTTPS link generating below. Copy and share it with anyone!)
echo.
echo ======================================================================

:: Launch SSH tunnel via localhost.run or pinggy
ssh -o StrictHostKeyChecking=no -p 443 -R0:localhost:5000 a.pinggy.io

pause
