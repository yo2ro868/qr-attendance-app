@echo off
rem Send one test mail with Notes (start and sign in to Notes first)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0send_notices.ps1" -AskTest
pause
