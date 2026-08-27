@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=C:\Users\Azi\miniconda3\envs\HZPBC\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

echo Building executable...
"%PYTHON_EXE%" "%~dp0tools\build_exe.py"
if errorlevel 1 (
    echo.
    echo Build failed. Please check the error shown above.
    if "%CI%"=="" pause
    exit /b 1
)

echo.
echo Build completed. See the dist folder.
if "%CI%"=="" pause
