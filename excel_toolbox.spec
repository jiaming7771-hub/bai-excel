# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

hiddenimports = (
    collect_submodules("pandas")
    + collect_submodules("openpyxl")
    + collect_submodules("xlrd")
    + collect_submodules("PySide6")
)

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=collect_data_files("openpyxl"),
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Excel小工具箱",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

app = BUNDLE(
    exe,
    name="Excel小工具箱.app",
    icon=None,
    bundle_identifier="local.excel.toolbox",
)
