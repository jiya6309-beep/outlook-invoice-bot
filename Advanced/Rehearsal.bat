@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo SHADOW RUN: reads your real mailboxes but changes NOTHING. A report is written to the Reports folder.
%PY% run_bot.py --shadow %*
echo.
pause
