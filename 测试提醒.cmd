@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0notify.ps1" -Force -NoState -Mode window -KeepSeconds 10
