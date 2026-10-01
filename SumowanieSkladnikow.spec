# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo,
    FixedFileInfo,
    StringFileInfo,
    StringTable,
    StringStruct,
    VarFileInfo,
    VarStruct,
)

version_info = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(1, 0, 0, 0),
        prodvers=(1, 0, 0, 0),
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo([
            StringTable(
                "040904B0",
                [
                    StringStruct("CompanyName", "SumowanieSkladnikow"),
                    StringStruct(
                        "FileDescription",
                        "Program do sumowania skladnikow Excel",
                    ),
                    StringStruct("FileVersion", "1.0.0.0"),
                    StringStruct("InternalName", "SumowanieSkladnikow"),
                    StringStruct(
                        "OriginalFilename",
                        "SumowanieSkladnikow.exe",
                    ),
                    StringStruct(
                        "ProductName",
                        "Sumowanie Skladnikow Excel",
                    ),
                    StringStruct("ProductVersion", "1.0.0.0"),
                ],
            )
        ]),
        VarFileInfo([
            VarStruct(
                "Translation",
                [1033, 1200],
            )
        ]),
    ],
)

td2_datas, td2_binaries, td2_hiddenimports = collect_all(
    "tkinterdnd2"
)

a = Analysis(
    ["sumowanie_skladnikow.py"],
    pathex=[],
    binaries=td2_binaries,
    datas=td2_datas,
    hiddenimports=td2_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SumowanieSkladnikow",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_info,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="SumowanieSkladnikow",
)
