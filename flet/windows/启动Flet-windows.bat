@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "RUN_VENV=%~dp0.venv-windows-run"
set "PYTHON_EXE=%RUN_VENV%\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
  echo [1/2] Creating Windows runtime environment...
  py -3 -m venv "%RUN_VENV%" 2>nul
  if errorlevel 1 python -m venv "%RUN_VENV%"
)

if not exist "%PYTHON_EXE%" (
  echo Failed to create the runtime environment. Please install 64-bit Python 3.10 or newer.
  pause
  exit /b 1
)

if not exist "%RUN_VENV%\.dependencies-ready" (
  echo [2/2] Installing Windows runtime dependencies...
  "%PYTHON_EXE%" -m pip install -r requirements-windows.txt
  if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
  )
  type nul > "%RUN_VENV%\.dependencies-ready"
)

"%PYTHON_EXE%" run.py --flet
if errorlevel 1 pause
