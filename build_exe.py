#!/usr/bin/env python3
"""
打包脚本：将 main.py 打包为独立 .exe（Windows）
在 Windows 上运行此脚本，或使用下方命令手动打包。
"""

BUILD_CMD = """
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "Excel批量表单生成器" ^
  --icon app.ico ^
  main.py
"""

INSTALL_CMD = "pip install pyinstaller openpyxl pandas"

print("=" * 60)
print("  Excel 批量表单生成器 — 打包指南")
print("=" * 60)
print()
print("【第一步】安装依赖（在 Windows 命令行运行）：")
print(f"  {INSTALL_CMD}")
print()
print("【第二步】打包为 .exe：")
print(BUILD_CMD)
print()
print("【说明】")
print("  --onefile   : 打包成单个 .exe 文件，无需安装 Python")
print("  --windowed  : 运行时不显示黑色控制台窗口")
print("  --icon      : 可选，设置 .ico 图标（如不需要可删除该行）")
print()
print("【打包完成后】")
print("  生成文件位于: dist/Excel批量表单生成器.exe")
print("  直接双击即可运行，无需安装任何环境！")
print("=" * 60)
