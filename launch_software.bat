@echo off
title The Indian Express - Story Count Software
color 0B
cls

:menu
cls
echo ==============================================================================
echo             THE INDIAN EXPRESS - STORY COUNT ^& ARTICLE ANALYZER
echo ==============================================================================
echo.
echo   [1] Run Interactive Software in Terminal (Black Screen Mode)
echo   [2] Launch Web Dashboard Software in Browser (http://localhost:5000)
echo   [3] Generate Public Live URL to Share with Team (100%% Cloudflare Immune)
echo   [4] Quick Analysis: Last 7 Days Story Count
echo   [5] Quick Analysis: Current Month to Date
echo   [6] Push Latest Updates to GitHub (Sync Streamlit Cloud)
echo   [7] Exit Software
echo.
echo ==============================================================================
set /p opt="Enter your choice [1-7]: "

if "%opt%"=="1" goto terminal_mode
if "%opt%"=="2" goto web_mode
if "%opt%"=="3" goto public_url
if "%opt%"=="4" goto last7days
if "%opt%"=="5" goto thismonth
if "%opt%"=="6" goto push_git
if "%opt%"=="7" goto exit_app

echo Invalid selection. Please enter a number between 1 and 7.
pause
goto menu

:terminal_mode
cls
echo ==============================================================================
echo              TERMINAL MODE - STORY COUNT ^& ARTICLE ANALYZER
echo ==============================================================================
echo.
cd /d "%~dp0"
python count_articles.py
echo.
echo Press any key to return to main menu...
pause >nul
goto menu

:web_mode
cls
echo ==============================================================================
echo             STARTING LOCAL WEB DASHBOARD SOFTWARE (PORT 5000)
echo ==============================================================================
echo.
echo  Starting local web server on http://localhost:5000 ...
echo  Your browser will open automatically.
echo  Press Ctrl+C in this window anytime to return to the menu.
echo.
cd /d "%~dp0"
python web_app.py
pause
goto menu

:public_url
cls
echo ==============================================================================
echo             GENERATING SHAREABLE PUBLIC HTTPS LIVE URL
echo ==============================================================================
echo.
call "%~dp0start_public_live_url.bat"
goto menu

:last7days
cls
echo ==============================================================================
echo                    QUICK ANALYSIS: LAST 7 DAYS
echo ==============================================================================
echo.
cd /d "%~dp0"
python -c "from datetime import datetime, timedelta; import count_articles; today = datetime.now().date(); count_articles.run_pipeline(str(today - timedelta(days=6)), str(today))"
echo.
echo Press any key to return to main menu...
pause >nul
goto menu

:thismonth
cls
echo ==============================================================================
echo                 QUICK ANALYSIS: CURRENT MONTH TO DATE
echo ==============================================================================
echo.
cd /d "%~dp0"
python -c "from datetime import datetime; import count_articles; today = datetime.now().date(); count_articles.run_pipeline(str(today.replace(day=1)), str(today))"
echo.
echo Press any key to return to main menu...
pause >nul
goto menu

:push_git
cls
echo ==============================================================================
echo                    SYNC WITH GITHUB ^& STREAMLIT CLOUD
echo ==============================================================================
echo.
cd /d "%~dp0"
git add .
git commit -m "Update story counter software and public URL launcher"
git push origin main
echo.
echo Code synced with GitHub! Streamlit Cloud will update automatically.
pause
goto menu

:exit_app
exit
