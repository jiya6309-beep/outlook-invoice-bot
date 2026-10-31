@echo off
rem Security run: reads the alarm reports of the Security mailbox and fills the fortnightly STORE OPEN ^& CLOSE Excel.
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
%PY% run_bot.py --security %*
pause
