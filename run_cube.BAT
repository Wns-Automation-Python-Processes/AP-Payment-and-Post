@echo off
REM Navigate to the directory where the script is located (optional if saved in the same folder)
cd /d "%~dp0"

REM Run the Python script
python cube.py

REM Keep the command prompt open after execution
pause