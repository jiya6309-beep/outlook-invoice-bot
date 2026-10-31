@echo off
title Outlook Invoice Bot
cd /d "%~dp0"
set "PY=python"
where python >nul 2>nul || set "PY=py -3"
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Python is not installed on this PC.
  echo Install it from https://www.python.org/downloads/ and TICK "Add python.exe to PATH" on the first screen.
  echo Then double-click this file again.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)
%PY% -c "import openpyxl, pdfplumber, PIL, pytesseract, win32com.client" >nul 2>nul
if errorlevel 1 (
  echo First start: installing the parts the bot needs. This takes about a minute and needs internet, once.
  %PY% -m pip install --disable-pip-version-check -r requirements.txt
  if errorlevel 1 (
    echo.
    echo The installation did not finish. Check the internet connection and double-click this file again.
    pause
    exit /b 1
  )
)
%PY% app.py
if errorlevel 1 pause
