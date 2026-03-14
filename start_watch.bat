@echo off
cd /d "%~dp0"
title VBA Watch
echo VBA Watch - vba_modules to Excel Import
echo.
python watch_and_export.py
if errorlevel 1 echo ERROR: Run "pip install pywin32 watchdog"
echo.
pause
