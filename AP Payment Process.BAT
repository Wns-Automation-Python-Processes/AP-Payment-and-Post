@echo off
title AP Payment Process
REM Navigate to the directory where the script is located
cd /d "%~dp0"

REM Launch the AP Payment Process UI using the project virtual environment
if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" sagex3_ui.py
) else (
    start "" pythonw sagex3_ui.py
)
