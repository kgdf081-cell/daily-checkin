@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo   正在安装「每日打卡」提醒任务...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup-task.ps1"
echo.
pause
