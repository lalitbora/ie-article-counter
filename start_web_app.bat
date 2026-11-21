@echo off
title The Indian Express - Article Counter Web Dashboard
color 0B
cls
echo ======================================================================
echo       THE INDIAN EXPRESS - ARTICLE COUNTER WEB DASHBOARD
echo ======================================================================
echo.
echo  Starting local web server on http://localhost:5000 ...
echo  Your browser will open automatically.
echo.
echo  Press Ctrl+C to stop the server anytime.
echo ======================================================================
echo.
cd /d "%~dp0"
python web_app.py
pause
