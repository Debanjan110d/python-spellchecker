@echo off
echo System-wide Spell Checker
echo ========================
echo.
echo Starting spell checker...
echo For best results, run as Administrator
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH
    echo Please install Python first
    pause
    exit /b 1
)

REM Run the spell checker
python main.py

echo.
echo Spell checker has stopped.
pause
