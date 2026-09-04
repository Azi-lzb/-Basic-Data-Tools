@echo off
rem Build the unified Flask shell EXE for Windows 7 (Python 3.7 + PyInstaller 4.10
rem from the HZPBCwin7 conda env). Output: dist\BasicAudit_Flask_Win7.exe + dist\core\
cd /d "%~dp0"
set PY=C:\Users\Azi\miniconda3\envs\HZPBCwin7\python.exe
if not exist "%PY%" (
  echo [ERROR] HZPBCwin7 env not found: %PY%
  pause
  exit /b 1
)
"%PY%" -c "import PyInstaller" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] PyInstaller missing in HZPBCwin7. Run: %PY% -m pip install pyinstaller==4.10
  pause
  exit /b 1
)
rem Flask 2.2.5 is the last line supporting Python 3.7.
"%PY%" -c "import flask" >nul 2>nul || "%PY%" -m pip install "flask==2.2.5" "werkzeug==2.2.3"
if errorlevel 1 (
  echo [ERROR] Failed to install flask 2.2.5 into HZPBCwin7.
  pause
  exit /b 1
)
"%PY%" build_exe.py --win7
if errorlevel 1 pause
