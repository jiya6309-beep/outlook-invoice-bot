@echo off
cd /d "%~dp0.."
echo This creates a Windows scheduled task that runs the bot every 30 minutes while YOU ARE SIGNED IN.
echo (Outlook must be open and signed in - Windows cannot run Outlook automation on a locked-out account.)
echo.
schtasks /Create /TN "Outlook Invoice Bot" /TR "\"%CD%\run_bot.bat\"" /SC MINUTE /MO 30 /F
echo.
echo To remove it later:  schtasks /Delete /TN "Outlook Invoice Bot" /F
pause
