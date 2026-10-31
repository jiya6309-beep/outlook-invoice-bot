@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo Drag your site list (.xlsx) onto this file, or type its path.
set "LIST=%~1"
if "%LIST%"=="" set /p LIST=Path of the NEW site list: 
%PY% import_sites.py "%LIST%"
echo.
pause
