@echo off
cd /d "%~dp0"
rem pywebview launcher (unified core). Launch the packaged EXE next to this
rem file if one exists, otherwise run the source entry point.
for %%E in ("%~dp0*.exe") do (
    start "" "%%~fE"
    exit /b 0
)
set PY=
python -c "import sys" >nul 2>nul && set PY=python
if not defined PY py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY (
  echo [ERROR] Python not found. Install Python and add it to PATH.
  pause
  exit /b 1
)
%PY% -c "import webview" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] pywebview missing. Run: pip install pywebview
  pause
  exit /b 1
)
%PY% run.py %*
if errorlevel 1 pause
