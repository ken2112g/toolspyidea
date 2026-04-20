@echo off
chcp 65001 >nul
title Tool Spy Idea - Setup
color 0A

echo.
echo  ╔═══════════════════════════════════════╗
echo  ║     Tool Spy Idea v1.0.0 - SETUP     ║
echo  ║          by ChThanh                   ║
echo  ╚═══════════════════════════════════════╝
echo.

:: Check Python
echo [1/3] Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ❌ Python chưa được cài!
    echo.
    echo  Tải Python tại: https://www.python.org/downloads/
    echo  QUAN TRỌNG: Tick ✅ "Add Python to PATH" khi cài!
    echo.
    start https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=2" %%a in ('python --version 2^>^&1') do echo     Python %%a ✓

:: Install packages
echo.
echo [2/3] Installing packages...
echo     Flask, Playwright, BeautifulSoup4, openpyxl, requests, Pillow, dropbox...
pip install flask playwright beautifulsoup4 openpyxl requests Pillow dropbox -q
if %errorlevel% neq 0 (
    echo  ❌ Lỗi cài packages! Thử chạy lại với quyền Admin.
    pause
    exit /b 1
)
echo     Packages ✓

:: Install Playwright browsers
echo.
echo [3/3] Installing Chrome for scraping...
echo     (Lần đầu mất 1-2 phút, lần sau không cần)
python -m playwright install chromium
if %errorlevel% neq 0 (
    echo  ⚠ Playwright chromium install failed, sẽ dùng Chrome của bạn thay thế.
)
echo     Browser ✓

echo.
echo  ╔═══════════════════════════════════════╗
echo  ║          ✅ SETUP HOÀN TẤT!          ║
echo  ║                                       ║
echo  ║  Double-click "ToolSpyIdea.bat"       ║
echo  ║  hoặc "START.vbs" để chạy app!       ║
echo  ╚═══════════════════════════════════════╝
echo.
pause
