@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo DEMO: sample e-mails in a simulated Outlook. Your real mailboxes are NOT touched.
%PY% run_bot.py --demo
echo.
pause
