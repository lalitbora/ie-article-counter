@echo off
title The Indian Express - Story Counter
color 0A
cls
echo ======================================================================
echo           THE INDIAN EXPRESS - DAILY ARTICLE ^& STORY COUNTER
echo ======================================================================
echo.
cd /d "%~dp0"
python count_articles.py
echo.
pause
