@echo off
chcp 65001 >nul
title Tool Spy Idea v1.0.0
color 0B
cd /d "%~dp0"

echo.
echo  ╔═══════════════════════════════════════╗
echo  ║       Tool Spy Idea v1.0.0           ║
echo  ║       by ChThanh                      ║
echo  ╠═══════════════════════════════════════╣
echo  ║  App đang chạy tại:                  ║
echo  ║  http://127.0.0.1:5123               ║
echo  ║                                       ║
echo  ║  Đóng cửa sổ này = tắt app           ║
echo  ╚═══════════════════════════════════════╝
echo.

python main.py
if %errorlevel% neq 0 (
    echo.
    echo  ❌ Lỗi! Chạy SETUP.bat trước nếu chưa setup.
    pause
)
