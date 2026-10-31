@echo off
rem Used by the scheduled task (no pause, so it never waits for a key).
cd /d "%~dp0"
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
%PY% run_bot.py %*
