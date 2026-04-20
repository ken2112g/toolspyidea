@echo off
chcp 65001 >nul
title Tool Spy Idea - Build EXE
color 0E
cd /d "%~dp0"

echo.
echo  ========================================
echo    Tool Spy Idea - BUILD .EXE
echo  ========================================
echo.

echo [1/3] Installing packages...
pip install pyinstaller flask playwright beautifulsoup4 openpyxl requests Pillow dropbox -q
python -m playwright install chromium
echo     OK

echo.
echo [2/3] Creating icon...
python build_helpers.py icon

echo.
echo [3/3] Building .exe (2-5 min)...
for /f "delims=" %%i in ('python build_helpers.py pw_path') do set PW_PATH=%%i

python -m PyInstaller --name "ToolSpyIdea" --noconfirm --clean --windowed --add-data "static;static" --add-data "data;data" --add-data "modules;modules" --add-data "extension;extension" --add-data "%PW_PATH%;playwright" --hidden-import "playwright.sync_api" --hidden-import "playwright._impl" --collect-all "playwright" --icon "app_icon.ico" main.py

xcopy /s /y /i data dist\ToolSpyIdea\data >nul 2>&1
xcopy /s /y /i static dist\ToolSpyIdea\static >nul 2>&1
xcopy /s /y /i extension dist\ToolSpyIdea\extension >nul 2>&1

:: Copy Playwright browser binaries
echo     Copying browser binaries...
for /f "delims=" %%i in ('python -c "from pathlib import Path; import playwright; d=Path(playwright.__file__).parent/'driver'/'package'/'.local-browsers'; print(d) if d.exists() else print('')"') do set PW_BROWSERS=%%i
if not "%PW_BROWSERS%"=="" (
    if exist "%PW_BROWSERS%" (
        xcopy /s /y /i "%PW_BROWSERS%" "dist\ToolSpyIdea\_internal\playwright\driver\package\.local-browsers" >nul 2>&1
        echo     Browser binaries copied OK
    )
)

echo.
echo  Done! File tai: dist\ToolSpyIdea\ToolSpyIdea.exe
explorer dist\ToolSpyIdea
pause
