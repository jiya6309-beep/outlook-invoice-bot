@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo Test the bot on a folder of SAVED .msg e-mails. Your real Outlook is NOT touched.
echo Drag the folder onto this file, or type its path.
set "MSGDIR=%~1"
if "%MSGDIR%"=="" set /p MSGDIR=Folder with the .msg files: 
%PY% run_bot.py --msg-folder "%MSGDIR%"
echo.
pause
