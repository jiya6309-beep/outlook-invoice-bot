@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo Read e-mails that ARRIVED between two dates (both days included).
set /p FROMD=From date (YYYY-MM-DD): 
set /p TOD=To date   (YYYY-MM-DD): 
set /p SHADOW=Rehearsal only, change nothing? (y/n): 
if /i "%SHADOW%"=="y" (%PY% run_bot.py --shadow --from %FROMD% --to %TOD%) else (%PY% run_bot.py --from %FROMD% --to %TOD%)
echo.
pause
