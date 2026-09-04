@echo off
cd /d "%~dp0"
rem Launch the packaged EXE next to this file if one exists, otherwise run
rem the source entry point with the current Python interpreter.
for %%E in ("%~dp0*.exe") do (
    start "" "%%~fE"
    exit /b 0
)
python run.py
if errorlevel 1 pause
