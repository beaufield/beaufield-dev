@echo off
python "%~dp0convert.py"
if errorlevel 1 (
    pause
    exit /b 1
)
pause
