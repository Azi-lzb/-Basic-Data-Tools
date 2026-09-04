@echo off
rem Build the unified Flask shell EXE (modern Windows, Python 3.8+).
rem Output: dist\BasicAudit_Flask.exe + dist\core\
cd /d "%~dp0"
set PY=
python -c "import sys" >nul 2>nul && set PY=python
if not defined PY py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY (
  echo [ERROR] Python not found. Install Python and add it to PATH.
  pause
  exit /b 1
)
%PY% -c "import flask, PyInstaller" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] flask / PyInstaller missing. Run: pip install flask openpyxl pywin32 pyinstaller
  pause
  exit /b 1
)
%PY% build_exe.py
if errorlevel 1 pause
