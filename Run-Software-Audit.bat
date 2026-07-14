@echo off
setlocal
cd /d "%~dp0"
start "" powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SoftwareAudit.ps1"
