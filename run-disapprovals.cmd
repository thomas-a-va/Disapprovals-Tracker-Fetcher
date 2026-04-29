@echo off
set AR_EMAIL=
set AR_PASSWORD=
set AR_CHROME_PROFILE=C:\PPRChrome
rem Session client (logged-in "default" context). Omit to use first client in clients.csv.
set AR_CURRENT_CLIENT_ID=2034
rem If version detection fails, set AR_VERSION to the value shown in the app (e.g. 2026/02/13 10:38:14 PST):
rem set AR_VERSION=2026/02/13 10:38:14 PST

set AR_SNAP_MODE=none

rem Apps Script webhook URL – creates sheets and moves them to shared drive.
rem Leave empty to skip Google Sheets export (CSV-only mode).
set AR_APPS_SCRIPT_URL=https://script.google.com/macros/s/AKfycbx28xfEu3-B1VoenBOJaMP1gQ115QdUqq5jVQK9XNZ1nsT4buufBtQkpeqXsVKbl_tJ/exec

node "%~dp0disapprovals.js"
if "%SCHEDULED%"=="" pause
