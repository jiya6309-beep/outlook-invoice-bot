@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo Installing the Python packages the bot needs ...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
echo.
echo Done. Next: 1_check_setup.bat
pause
