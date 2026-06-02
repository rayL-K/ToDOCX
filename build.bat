@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "LOG=%TEMP%\todocx-build.log"

echo [1/4] Cleaning old builds...
if exist build rmdir /s /q build >nul 2>nul
if exist dist\ToDOCX rmdir /s /q dist\ToDOCX >nul 2>nul
echo [1/4] Done>>"%LOG%"

echo [2/4] Checking spec config...
python -c "p=\'ToDOCX.spec\';c=open(p,\'r\',encoding=\'utf-8\').read();m=(\'templates\',\'templates\');print(\'ok\' if m in c else \'missing\')"
if %errorlevel% neq 0 (
    echo [2/4] Failed>>"%LOG%"
    pause
    exit /b 1
)
echo [2/4] Done>>"%LOG%"

echo [3/4] Running PyInstaller...
pyinstaller ToDOCX.spec --noconfirm
if %errorlevel% neq 0 (
    echo [3/4] Failed>>"%LOG%"
    pause
    exit /b 1
)
echo [3/4] Done>>"%LOG%"

echo [4/4] Creating zip archive...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\ToDOCX\*' -DestinationPath 'dist\ToDOCX-v1.0.4-win64.zip' -Force"
if %errorlevel% neq 0 (
    echo [4/4] Failed>>"%LOG%"
    pause
    exit /b 1
)
echo [4/4] Done>>"%LOG%"

echo.
echo All done. Output: dist\ToDOCX\
echo Log: %LOG%
pause