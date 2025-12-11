"""生成 PyInstaller 版本信息文件"""

# 从 src.version 导入版本信息
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.version import __version__, __author__, __app_name__, VERSION_INFO, get_version_tuple

VERSION_TUPLE = get_version_tuple()

# PyInstaller 版本信息模板
VERSION_INFO_TEMPLATE = f"""
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={VERSION_TUPLE},
    prodvers={VERSION_TUPLE},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          u'080404b0',
          [
            StringStruct(u'CompanyName', u'{__author__}'),
            StringStruct(u'FileDescription', u'{VERSION_INFO["description"]}'),
            StringStruct(u'FileVersion', u'{__version__}'),
            StringStruct(u'InternalName', u'{__app_name__}'),
            StringStruct(u'LegalCopyright', u'{VERSION_INFO["copyright"]}'),
            StringStruct(u'OriginalFilename', u'{__app_name__}.exe'),
            StringStruct(u'ProductName', u'{__app_name__}'),
            StringStruct(u'ProductVersion', u'{__version__}'),
          ]
        )
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)
"""

if __name__ == "__main__":
    # 生成版本信息文件
    output_path = os.path.join(os.path.dirname(__file__), "file_version_info.txt")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(VERSION_INFO_TEMPLATE.strip())
    print(f"版本信息已生成: {output_path}")
    print(f"版本号: {__version__}")
