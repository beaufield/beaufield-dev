@echo off
python "%~dp0convert_to_import.py"
if errorlevel 1 (
    pause
    exit /b 1
)
pause
