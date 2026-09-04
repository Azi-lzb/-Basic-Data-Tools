@echo off
rem Build the unified pywebview shell EXE (modern Windows, WebView2).
rem Output: dist\<exe>.exe + dist\core\
cd /d "%~dp0"
set PY=
python -c "import sys" >nul 2>nul && set PY=python
if not defined PY py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY (
  echo [ERROR] Python not found. Install Python and add it to PATH.
  pause
  exit /b 1
)
%PY% -c "import webview, PyInstaller" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] pywebview / PyInstaller missing. Run: pip install pywebview openpyxl pywin32 pyinstaller
  pause
  exit /b 1
)
%PY% build_exe.py
if errorlevel 1 pause
