@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   ToDOCX 打包脚本
echo ============================================
echo.

REM ----- 清理旧构建 -----
echo [1/4] 清理旧构建...
rmdir /s /q build >nul 2>nul
rmdir /s /q dist\ToDOCX >nul 2>nul
echo   OK

REM ----- 更新 spec 中的 datas（把 templates 目录打包进去）-----
echo [2/4] 更新打包配置...
python -c "
import json
# 读取现有 spec，追加 templates 目录到 datas
spec_path = 'ToDOCX.spec'
with open(spec_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 在 datas 中追加 templates 目录（如果还没有的话）
if \"('templates', 'templates')\" not in content:
    content = content.replace(
        \"datas=[('src/ui/docx.ico', 'src/ui')]\",
        \"datas=[('src/ui/docx.ico', 'src/ui'), ('templates', 'templates')]\",
    )
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('   已追加 templates 目录到打包配置')
else:
    print('   打包配置已包含 templates 目录，跳过')
"
echo   OK

REM ----- 执行 PyInstaller -----
echo [3/4] 执行 PyInstaller...
pyinstaller ToDOCX.spec --noconfirm
echo   OK

REM ----- 打包为 zip -----
echo [4/4] 打包为 zip...
set "VERSION=1.0.4"
set "ZIP_NAME=ToDOCX-v%VERSION%-win64.zip"
powershell -NoProfile -Command ^
    \"Compress-Archive -Path 'dist\ToDOCX\*' -DestinationPath 'dist\%ZIP_NAME%' -Force\"
echo   OK

echo.
echo ============================================
echo   打包完成！
echo   输出目录: dist\ToDOCX\
echo   ZIP 文件: dist\%ZIP_NAME%
echo ============================================
pause
