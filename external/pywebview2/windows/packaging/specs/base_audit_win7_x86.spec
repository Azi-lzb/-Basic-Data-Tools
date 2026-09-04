# -*- mode: python ; coding: utf-8 -*-
"""ASCII-named 32-bit Windows 7 spec. Build only with HZPBCwin7x86 (Python 3.7)."""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


SPEC_DIR = Path(SPECPATH).resolve()
PROJECT_ROOT = next(
    parent for parent in (SPEC_DIR, *SPEC_DIR.parents)
    if (parent / "run.py").is_file()
)
webview_data = collect_data_files(
    "webview",
    includes=[
        "js/*.js",
        "js/lib/*.js",
        "lib/Microsoft.Web.WebView2.Core.dll",
        "lib/Microsoft.Web.WebView2.WinForms.dll",
        "lib/WebBrowserInterop.x86.dll",
        "lib/runtimes/win-arm64/native/WebView2Loader.dll",
        "lib/runtimes/win-x64/native/WebView2Loader.dll",
        "lib/runtimes/win-x86/native/WebView2Loader.dll",
    ],
)

# Keep pythonnet's CLR bridge beside its package after PyInstaller unpacks the
# one-file EXE.  Without it pywebview cannot initialise the WinForms backend.
pythonnet_data = collect_data_files(
    "pythonnet",
    includes=[
        "runtime/Python.Runtime.dll",
        "runtime/Python.Runtime.deps.json",
    ],
)

a = Analysis(
    [str(PROJECT_ROOT / "run.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=[(str(PROJECT_ROOT / "frontend" / "web"), "web")] + webview_data + pythonnet_data,
    hiddenimports=[
        "base_audit.excel_com", "base_audit.service", "base_audit.web_app",
        "base_audit.region_summary", "pythoncom", "pywintypes", "win32com",
        "win32com.client", "openpyxl", "xlrd", "webview", "webview.guilib",
        "webview.platforms.winforms", "webview.platforms.edgechromium", "clr",
        "pythonnet", "clr_loader",
    ],
    hookspath=[str(PROJECT_ROOT / "packaging" / "hooks" / "win7_x86")],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "doctest", "pydoc", "pytest", "unittest", "tkinter", "numpy", "pandas",
        "lxml", "lxml.etree", "lxml.objectify", "lxml.html", "lxml.isoschematron",
        "webview.platforms.android", "webview.platforms.cocoa", "webview.platforms.gtk",
        "webview.platforms.qt", "webview.platforms.cef", "webview.platforms.mshtml",
        "pythonwin", "win32ui",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
    # See the x64 spec: use an ASCII staging name for PyInstaller 4 / pefile.
    name="BaseAuditTool_Win7_x86",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
