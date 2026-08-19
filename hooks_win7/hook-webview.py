"""Minimal pywebview hook for the x64 Windows 7 build.

The upstream hook collects every platform bridge.  This application always
uses the Windows Edge Chromium backend, so retain only the four files it loads
at runtime.
"""

from PyInstaller.utils.hooks import collect_data_files


datas = collect_data_files(
    "webview",
    includes=[
        # pywebview injects these scripts into the local HTML page to create
        # window.pywebview.api and fire the pywebviewready event.
        "js/*.js",
        "js/lib/*.js",
        "lib/Microsoft.Web.WebView2.Core.dll",
        "lib/Microsoft.Web.WebView2.WinForms.dll",
        "lib/WebBrowserInterop.x64.dll",
        # pywebview 5 probes all three directories during Edge startup.  The
        # x86/arm64 loaders are not executed by this x64 package, but their
        # presence prevents the probe from aborting before it reaches x64.
        "lib/runtimes/win-arm64/native/WebView2Loader.dll",
        "lib/runtimes/win-x64/native/WebView2Loader.dll",
        "lib/runtimes/win-x86/native/WebView2Loader.dll",
    ],
)
