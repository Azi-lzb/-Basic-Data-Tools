@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem Build the Win7 x64 package with the HZPBCwin7 conda env.
rem ASCII-only + CRLF on purpose: batch parsing must not depend on the
rem console code page, and cmd.exe misparses LF-only batch files.

pushd "%~dp0"
if errorlevel 1 goto :folder_error

set "PYINSTALLER_CONFIG_DIR=%CD%\build\win7_cache"
set "CONDA_EXE=%UserProfile%\miniconda3\Scripts\conda.exe"

rem Locate the x64 spec by wildcard so no Chinese filename appears here.
for %%S in ("%CD%\*_win7.spec") do set "SPEC_FILE=%%~fS"
if not exist "%CONDA_EXE%" goto :conda_error
if not defined SPEC_FILE goto :spec_error

echo.
echo Building Win7 x64 package with HZPBCwin7 ...
"%CONDA_EXE%" run -n HZPBCwin7 python -m PyInstaller --clean --noconfirm --distpath "dist" --workpath "build\win7" "%SPEC_FILE%"
if errorlevel 1 goto :build_error

for %%O in ("%CD%\dist\*_Win7.exe") do set "OUTPUT_FILE=%%~fO"
if not defined OUTPUT_FILE goto :output_error
if not exist "dist\data" mkdir "dist\data"
if not exist "dist\data\config.xlsx" copy /y "data\config.xlsx" "dist\data\config.xlsx" >nul
for %%F in (*.docx) do copy /y "%%~fF" "dist\" >nul

echo.
echo Build completed. See dist\*_Win7.exe
goto :end_ok

:folder_error
echo ERROR: Cannot enter the project folder.
goto :end_fail
:conda_error
echo ERROR: Conda was not found at %CONDA_EXE%
goto :end_fail
:spec_error
echo ERROR: Win7 x64 spec file was not found.
goto :end_fail
:build_error
echo ERROR: Win7 x64 package build failed. Read the messages above.
goto :end_fail
:output_error
echo ERROR: EXE was not generated.
goto :end_fail
:end_ok
popd
echo.
pause
exit /b 0
:end_fail
popd
echo.
pause
exit /b 1
