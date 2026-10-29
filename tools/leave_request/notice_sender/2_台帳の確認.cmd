@echo off
rem Read the synced ledger and list the notices that would be sent (sends nothing)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0send_notices.ps1" -DryRun -Once
pause
