# -*- mode: python ; coding: utf-8 -*-
"""Windows 7 compatible build: Python 3.7 + PyInstaller 4.10.

Build this spec only from the HZPBCwin7 conda environment.  Keep the local
HTML workbench and the x64 WebView2 bridge, but exclude unused UI backends and
development packages to control the single-file executable size.
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


PROJECT_ROOT = Path(SPECPATH).resolve()

# pywebview's Edge Chromium backend locates these files relative to its module.
# This is deliberately x64-only, matching the Win7 build interpreter.
webview_data = collect_data_files(
    "webview",
    includes=[
        "js/*.js",
        "js/lib/*.js",
        "lib/Microsoft.Web.WebView2.Core.dll",
        "lib/Microsoft.Web.WebView2.WinForms.dll",
        "lib/WebBrowserInterop.x64.dll",
        "lib/runtimes/win-arm64/native/WebView2Loader.dll",
        "lib/runtimes/win-x64/native/WebView2Loader.dll",
        "lib/runtimes/win-x86/native/WebView2Loader.dll",
    ],
)

a = Analysis(
    [str(PROJECT_ROOT / "run.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=[(str(PROJECT_ROOT / "src" / "base_audit" / "web"), "web")] + webview_data,
    hiddenimports=[
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
        "webview.guilib",
        # pywebview 5 first loads WinForms, which then selects Edge Chromium.
        "webview.platforms.winforms",
        "webview.platforms.edgechromium",
        "clr",
        "pythonnet",
        "clr_loader",
    ],
    hookspath=[str(PROJECT_ROOT / "hooks_win7")],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "doctest", "pydoc", "pytest", "unittest", "tkinter",
        "numpy", "pandas", "lxml", "lxml.etree", "lxml.objectify",
        "lxml.html", "lxml.isoschematron",
        "webview.platforms.android", "webview.platforms.cocoa",
        "webview.platforms.gtk", "webview.platforms.qt",
        "webview.platforms.cef", "webview.platforms.mshtml",
        "pythonwin", "win32ui",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="基础数据审核工具_Win7",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
