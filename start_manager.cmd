@echo off
cd /d "%~dp0"
title Excel Sync Manager
python app.py
echo.
echo Manager exited with code %errorlevel%.
pause
