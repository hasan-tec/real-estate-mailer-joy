@echo off
REM Real Estate Mailer Generator Launcher
REM This batch file activates the virtual environment and runs the GUI app

echo ========================================
echo Real Estate Mailer Generator
echo ========================================
echo.

REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "env\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found!
    echo Please ensure the 'env' folder exists in the same directory as this batch file.
    echo.
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call env\Scripts\activate.bat

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in virtual environment!
    echo.
    pause
    exit /b 1
)

REM Run the GUI application
echo Starting Mailer Generator GUI...
echo.
python mailer_app.py

REM If the app exits with an error
if errorlevel 1 (
    echo.
    echo ========================================
    echo Application exited with an error.
    echo ========================================
    pause
)

REM Deactivate virtual environment (optional, as the window will close anyway)
deactivate
