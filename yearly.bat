@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo Python was not found on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during installation.
    echo.
    pause
    exit /b 1
)

python -m pip install --quiet --disable-pip-version-check openpyxl flask
if errorlevel 1 (
    echo Failed to install dependencies.
    pause
    exit /b 1
)

set "FOLDER=%~1"
if "%FOLDER%"=="" set "FOLDER=payments"

if not exist "%FOLDER%" (
    echo Creating folder "%FOLDER%"...
    mkdir "%FOLDER%"
    echo.
    echo Drop your monthly .xlsx files into "%FOLDER%" and they will be
    echo processed automatically. Press Ctrl-C to stop watching.
    echo.
)

python payments_yearly.py "%FOLDER%" --watch --open

pause
