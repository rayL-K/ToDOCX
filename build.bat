@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   ToDOCX Build Script
echo ============================================
echo.

REM ----- Clean old builds -----
echo [1/4] Cleaning old builds...
if exist build rmdir /s /q build >nul 2>nul
if exist dist\ToDOCX rmdir /s /q dist\ToDOCX >nul 2>nul
echo   OK

REM ----- Update spec to include templates -----
echo [2/4] Updating build config...
python -c "import sys; sys.path.insert(0,'.'); p='ToDOCX.spec'; c=open(p,'r',encoding='utf-8').read(); exec(c.replace('datas=[(\"src/ui/docx.ico\",\"src/ui\")]','#keep'),{'__builtins__':__builtins__})" 2>nul || (
    python -c "
p='ToDOCX.spec'
c=open(p,'r',encoding='utf-8').read()
if 'templates' not in c:
    c=c.replace(\"datas=[('src/ui/docx.ico', 'src/ui')]\",\"datas=[('src/ui/docx.ico', 'src/ui'), ('templates', 'templates')]\")
    open(p,'w',encoding='utf-8').write(c)
    print('  templates added to spec')
else:
    print('  templates already in spec')
"
)
echo   OK

REM ----- Run PyInstaller -----
echo [3/4] Running PyInstaller...
pyinstaller ToDOCX.spec --noconfirm
echo   OK

REM ----- Create zip archive -----
echo [4/4] Creating zip archive...
set "VERSION=1.0.4"
set "ZIP_NAME=ToDOCX-v%VERSION%-win64.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\ToDOCX\*' -DestinationPath 'dist\%ZIP_NAME%' -Force"
echo   OK

echo.
echo ============================================
echo   Build complete!
echo   Output: dist\ToDOCX\
echo   Zip:    dist\%ZIP_NAME%
echo ============================================
pause
