$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root ".venv\Scripts\python.exe"
$pyinstaller = Join-Path $root ".venv\Scripts\pyinstaller.exe"
$distDir = Join-Path $root "dist"
$buildDir = Join-Path $root "build"

if (-not (Test-Path $python)) {
    throw "未找到虚拟环境 Python：$python"
}

if (-not (Test-Path $pyinstaller)) {
    throw "未找到 PyInstaller：$pyinstaller"
}

Push-Location $root
try {
    & $python "version_info.py"
    if (Test-Path $buildDir) {
        Remove-Item -Recurse -Force $buildDir
    }
    if (Test-Path $distDir) {
        Remove-Item -Recurse -Force $distDir
    }
    & $pyinstaller "ToDOCX.spec" "--clean" "--noconfirm"
}
finally {
    Pop-Location
}
