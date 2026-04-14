@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo Python was not found on PATH.
    echo Install Python 3.10 or newer from https://www.python.org/downloads/
    echo and make sure you tick "Add python.exe to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Installing/updating dependencies (openpyxl, flask)...
python -m pip install --quiet --disable-pip-version-check openpyxl flask
if errorlevel 1 (
    echo.
    echo Failed to install dependencies. Check your internet connection.
    pause
    exit /b 1
)

echo.
python payments_server.py --open

pause
