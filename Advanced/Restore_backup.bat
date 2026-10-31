@echo off
cd /d "%~dp0.."
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
echo These are the database backups. To restore one, run this again with its number, e.g.  Restore_backup.bat 1
echo (Close the bot first. The current database is kept next to it so the restore can be undone.)
echo.
%PY% restore_backup.py %*
echo.
pause
