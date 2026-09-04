@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

rem Windows Flet release.  The app uses Excel/WPS through pywin32 when the
rem corresponding office suite is installed on the user's computer.
set "BUILD_VENV=%~dp0.venv-win-build"
set "PYTHON_EXE=%BUILD_VENV%\Scripts\python.exe"
rem Flutter's Windows toolchain still mishandles some non-ASCII source paths.
rem Stage only the release source in a stable ASCII path; the original project
rem remains untouched and the finished folder is copied back on success.
set "STAGE_ROOT=C:\FletBuild"
set "STAGED_APP=%STAGE_ROOT%\source"
set "STAGED_OUTPUT=%STAGE_ROOT%\dist\windows"

if not exist "%PYTHON_EXE%" (
    echo [1/3] Creating isolated Windows build environment...
    py -3 -m venv "%BUILD_VENV%" 2>nul
    if errorlevel 1 python -m venv "%BUILD_VENV%"
)

if not exist "%PYTHON_EXE%" (
    echo Failed to create the build environment. Please install 64-bit Python 3.10 or newer.
    if not defined BASE_AUDIT_NO_PAUSE pause
    exit /b 1
)

echo [2/3] Installing minimal release dependencies...
"%PYTHON_EXE%" -m pip install --upgrade pip
if errorlevel 1 goto :failed
"%PYTHON_EXE%" -m pip install -r requirements-windows.txt
if errorlevel 1 goto :failed

echo [3/3] Building Windows release...
:stage_source
echo Preparing ASCII-path build source...
robocopy "%~dp0" "%STAGED_APP%" /MIR /XD ".venv-win-build" "build" "dist" "engine_trials" "tests" ".claude" /XF "windows-build*.log"
if errorlevel 8 goto :failed
echo Building from: %STAGED_APP%

:build_retry
"%BUILD_VENV%\Scripts\flet.exe" build windows "%STAGED_APP%" --yes --python-version "3.12" --no-rich-output --output "%STAGED_OUTPUT%" --exclude build --exclude .venv-win-build --exclude engine_trials --exclude tests --project "base-audit-windows" --product "基础数据审核工具（Windows版）" --artifact "基础数据审核工具-Windows" --cleanup-app --cleanup-packages --no-compile-packages
if not errorlevel 1 goto :built

if not defined BUILD_ATTEMPT set BUILD_ATTEMPT=1
if "%BUILD_ATTEMPT%"=="3" goto :failed
set /a BUILD_ATTEMPT+=1
echo Flutter toolchain download/build failed; retrying (%BUILD_ATTEMPT%/3)...
goto :build_retry

:built
robocopy "%STAGED_OUTPUT%" "%~dp0dist\windows" /MIR >nul
if errorlevel 8 goto :failed
echo.
echo Build complete: %~dp0dist\windows
echo Note: the package includes Excel/WPS COM support. Microsoft Excel or WPS itself must be installed on the target computer.
if not defined BASE_AUDIT_NO_PAUSE pause
exit /b 0

:failed
echo.
echo Build failed. Review the output above.
if not defined BASE_AUDIT_NO_PAUSE pause
exit /b 1
