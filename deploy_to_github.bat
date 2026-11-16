@echo off
title Push to GitHub for Cloud Hosting
cls
echo ======================================================================
echo             GITHUB REPOSITORY SETUP FOR CLOUD HOSTING
echo ======================================================================
echo.
echo  Step 1: Browser mein ye link kholiye aur nayi empty repository banayein:
echo         https://github.com/new
echo.
echo  Step 2: Repository name daalein (jaise: ie-article-counter) aur 
echo         'Create repository' button par click karein.
echo.
echo ======================================================================
set /p REPO_URL="Step 3: Apni GitHub repo ka URL yahan paste karein: "
if "%REPO_URL%"=="" goto error

git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
git branch -M main
echo.
echo Pushing code to GitHub...
git push -u origin main
echo.
echo ======================================================================
echo [SUCCESS] Code GitHub par upload ho gaya hai!
echo.
echo Ab agla step:
echo 1. https://share.streamlit.io par jayein aur GitHub se login karein.
echo 2. 'New App' dabayein aur is repo ko select karke 'Deploy' karein!
echo ======================================================================
goto end

:error
echo [Error] URL enter nahi kiya gaya tha. Dobara try karein.

:end
pause
