# -*- mode: python ; coding: utf-8 -*-
# 打包: pyinstaller Excel工作流.spec

from PyInstaller.building.api import EXE, PYZ
from PyInstaller.building.build_main import Analysis
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all("NodeGraphQt")

block_cipher = None

a = Analysis(
    ["run_workflow.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports
    + [
        "excel_workflow",
        "excel_workflow.nodes",
        "excel_workflow.core",
        "excel_workflow.ui",
        "openpyxl",
        "pdfplumber",
        "watchdog",
        "apscheduler",
        "openai",
        "PIL",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Excel工作流",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
