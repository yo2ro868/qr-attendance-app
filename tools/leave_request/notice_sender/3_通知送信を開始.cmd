@echo off
rem Check the ledger every few minutes and send unsent notices with Notes
rem Closing this window stops sending (keep it minimized)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0send_notices.ps1"
pause
