# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('src/ui/docx.ico', 'src/ui')],
    hiddenimports=[
        'markdown',
        'markdown.extensions',
        'markdown.extensions.tables',
        'markdown.extensions.fenced_code',
        'bs4',
        'lxml',
        'lxml.etree',
        'PIL',
        'PIL.Image',
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'mammoth',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ToDOCX',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon=['src/ui/docx.ico'],
    version='file_version_info.txt',
)
