# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

PROJECT_ROOT = Path(SPECPATH).resolve()

a = Analysis(
    [str(PROJECT_ROOT / "run.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=[(str(PROJECT_ROOT / "src" / "base_audit" / "web"), "web")],
    hiddenimports=[
        # Core audit modules are listed explicitly so a prior syntax error or
        # an optional UI entry path cannot leave the single-file EXE incomplete.
        "base_audit.excel_com",
        "base_audit.service",
        "base_audit.web_app",
        "base_audit.region_summary",
        "pythoncom",
        "pywintypes",
        "win32com",
        "win32com.client",
        "openpyxl",
        "xlrd",
        "webview",
        "webview.platforms.edgechromium",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "doctest",
        "pydoc",
        "pytest",
        "unittest",
        "tkinter",
        "numpy",
        "pandas",
        # openpyxl can fall back to the Python XML implementation; the audit
        # tool does not use lxml's HTML, schema or object APIs.
        "lxml",
        "lxml.etree",
        "lxml.objectify",
        "lxml.html",
        "lxml.isoschematron",
        # The local workbench is fixed to Windows Edge WebView2.
        "webview.platforms.android",
        "webview.platforms.cocoa",
        "webview.platforms.gtk",
        "webview.platforms.qt",
        # Excel COM needs pywin32 core, not the legacy PythonWin UI layer.
        "pythonwin",
        "win32ui",
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="基础数据审核工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
