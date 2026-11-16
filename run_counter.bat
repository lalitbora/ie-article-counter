@echo off
title The Indian Express - Article Counter
cls
echo ======================================================================
echo           THE INDIAN EXPRESS - DAILY ARTICLE COUNTER
echo ======================================================================
echo.
cd /d "%~dp0"
python count_articles.py
echo.
pause
