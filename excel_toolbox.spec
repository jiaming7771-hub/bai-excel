# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=collect_data_files("openpyxl"),
    hiddenimports=["pandas", "openpyxl", "xlrd", "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "IPython", "notebook", "tkinter"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Excel小工具箱",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="Excel小工具箱",
)

app = BUNDLE(
    coll,
    name="Excel小工具箱.app",
    icon=None,
    bundle_identifier="local.excel.toolbox",
    info_plist={
        "CFBundleName": "Excel小工具箱",
        "CFBundleDisplayName": "Excel小工具箱",
        "NSHighResolutionCapable": True,
    },
)
