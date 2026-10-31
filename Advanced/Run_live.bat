@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo LIVE RUN. Outlook must be open. Close Bot_Settings.xlsx first.
%PY% run_bot.py %*
echo.
pause
