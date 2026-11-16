@echo off
title The Indian Express - Article Counter Web App
cls
echo ======================================================================
echo       THE INDIAN EXPRESS - ARTICLE COUNTER WEB DASHBOARD
echo ======================================================================
echo.
echo  Starting local web server on http://localhost:5000 ...
echo  Aapka browser automatically open ho jayega.
echo.
echo  Press Ctrl+C to stop the server anytime.
echo ======================================================================
echo.
cd /d "%~dp0"
python web_app.py
pause
