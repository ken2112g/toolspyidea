@echo off
chcp 65001 >nul
title Tool Spy Idea - Stop
echo.
echo  Đang tắt Tool Spy Idea...
taskkill /f /im python.exe /fi "WINDOWTITLE eq Tool Spy*" >nul 2>&1
:: Kill any python running main.py on port 5123
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5123 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1
echo  ✓ Đã tắt!
timeout /t 2 >nul
