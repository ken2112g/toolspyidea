@echo off
chcp 65001 >nul
title Tool Spy Idea - Build Installer
color 0E
cd /d "%~dp0"

echo.
echo  ========================================
echo    Tool Spy Idea - BUILD INSTALLER
echo    Tao file Setup_ToolSpyIdea.exe
echo  ========================================
echo.

:: STEP 1
echo [1/5] Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Python chua cai! Tai: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo     Python OK

:: STEP 2
echo.
echo [2/5] Installing packages...
pip install pyinstaller flask playwright beautifulsoup4 openpyxl requests Pillow dropbox -q
python -m playwright install chromium
echo     Packages OK

:: STEP 3
echo.
echo [3/5] Creating app icon...
python build_helpers.py icon
if %errorlevel% neq 0 (
    echo     Icon creation failed, continuing without icon...
)

:: STEP 4
echo.
echo [4/5] Building .exe with PyInstaller (2-5 min)...

for /f "delims=" %%i in ('python build_helpers.py pw_path') do set PW_PATH=%%i
echo     Playwright path: %PW_PATH%

python -m PyInstaller --name "ToolSpyIdea" --noconfirm --clean --windowed --add-data "static;static" --add-data "data;data" --add-data "modules;modules" --add-data "extension;extension" --add-data "%PW_PATH%;playwright" --hidden-import "playwright.sync_api" --hidden-import "playwright._impl" --collect-all "playwright" --icon "app_icon.ico" main.py

if %errorlevel% neq 0 (
    echo  PyInstaller FAILED!
    pause
    exit /b 1
)

xcopy /s /y /i data dist\ToolSpyIdea\data >nul 2>&1
xcopy /s /y /i static dist\ToolSpyIdea\static >nul 2>&1
xcopy /s /y /i extension dist\ToolSpyIdea\extension >nul 2>&1

:: Copy Playwright browser binaries
echo     Copying browser binaries...
python -c "import subprocess,os; r=subprocess.run(['python','-m','playwright','install','--dry-run','chromium'],capture_output=True,text=True); print(r.stdout); print(r.stderr)" >nul 2>&1
for /f "delims=" %%i in ('python -c "from pathlib import Path; import playwright; d=Path(playwright.__file__).parent/'driver'/'package'/'.local-browsers'; print(d) if d.exists() else print('')"') do set PW_BROWSERS=%%i
if not "%PW_BROWSERS%"=="" (
    if exist "%PW_BROWSERS%" (
        xcopy /s /y /i "%PW_BROWSERS%" "dist\ToolSpyIdea\_internal\playwright\driver\package\.local-browsers" >nul 2>&1
        echo     Browser binaries copied OK
    )
)
echo     .exe built OK

:: STEP 5
echo.
echo [5/5] Building installer...

python build_helpers.py wizard

set "INNO_PATH="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "INNO_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"

if "%INNO_PATH%"=="" (
    echo.
    echo  Inno Setup 6 chua cai!
    echo  Tai mien phi: https://jrsoftware.org/isdl.php
    echo  Cai xong chay lai BUILD_INSTALLER.bat
    echo.
    echo  HOAC: Dung folder dist\ToolSpyIdea\ truc tiep
    echo  Zip lai gui cho nguoi khac giai nen dung.
    echo.
    start https://jrsoftware.org/isdl.php
    explorer dist\ToolSpyIdea
    pause
    exit /b 0
)

"%INNO_PATH%" installer.iss

if %errorlevel% neq 0 (
    echo  Inno Setup compile FAILED!
    pause
    exit /b 1
)

echo.
echo  ========================================
echo    INSTALLER CREATED!
echo    File: installer_output\Setup_ToolSpyIdea_v1.0.0.exe
echo    Gui file nay cho ai cung cai duoc!
echo  ========================================
echo.
explorer installer_output
pause
