@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
%PY% check_setup.py
echo.
pause
