@echo off
rem Flask audit tool launcher (unified core). ASCII-only + CRLF: cmd parses
rem batch files in the ANSI codepage, so Chinese text or LF endings break the
rem script (flash-close).
cd /d "%~dp0"
set PY=
python -c "import sys" >nul 2>nul && set PY=python
if not defined PY py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY (
  echo [ERROR] Python not found. Install Python and add it to PATH.
  pause
  exit /b 1
)
%PY% -c "import flask" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] flask missing. Run: pip install -r requirements.txt
  pause
  exit /b 1
)
%PY% run.py %*
if errorlevel 1 pause
