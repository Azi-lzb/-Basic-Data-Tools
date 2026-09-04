"""Minimal pywebview hook for the 32-bit Windows 7 build."""

from PyInstaller.utils.hooks import collect_data_files


datas = collect_data_files(
    "webview",
    includes=[
        "js/*.js",
        "js/lib/*.js",
        "lib/Microsoft.Web.WebView2.Core.dll",
        "lib/Microsoft.Web.WebView2.WinForms.dll",
        "lib/WebBrowserInterop.x86.dll",
        # pywebview 5 probes all three runtime directories before selecting x86.
        "lib/runtimes/win-arm64/native/WebView2Loader.dll",
        "lib/runtimes/win-x64/native/WebView2Loader.dll",
        "lib/runtimes/win-x86/native/WebView2Loader.dll",
    ],
)
