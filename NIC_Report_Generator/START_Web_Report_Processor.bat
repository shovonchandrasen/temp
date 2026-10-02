@echo off
setlocal
chcp 65001 >nul
title NIC Report Generator

echo ============================================================
echo   NIC Report Generator
echo ============================================================
echo.

:: --- Change to the directory where this .bat lives ---
cd /d "%~dp0"

:: --- Check Python ---
echo [1/4] Checking for Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python not found in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo and tick "Add Python to PATH" during install.
    pause
    exit /b 1
)
for /f "delims=" %%v in ('python --version 2^>^&1') do echo       Found %%v

:: --- Install / check dependencies ---
echo.
echo [2/4] Checking required Python libraries...
python -c "import flask" >nul 2>nul
if %errorlevel% neq 0 (
    echo       Installing flask...
    pip install flask
) else ( echo       flask OK )

python -c "import pandas" >nul 2>nul
if %errorlevel% neq 0 (
    echo       Installing pandas...
    pip install pandas
) else ( echo       pandas OK )

python -c "import openpyxl" >nul 2>nul
if %errorlevel% neq 0 (
    echo       Installing openpyxl...
    pip install openpyxl
) else ( echo       openpyxl OK )

:: --- Make sure working folders exist ---
if not exist "webapp\uploads" mkdir "webapp\uploads"
if not exist "webapp\outputs" mkdir "webapp\outputs"

:: --- Launch the Flask server ---
echo.
echo [3/4] Starting the web server...
echo       Your browser will open automatically at http://127.0.0.1:5000/
echo       If it does not, open that URL manually.
echo.
echo [4/4] Keep this window open while you use NIC Report Generator.
echo       Press Ctrl+C here to stop the server.
echo ============================================================
echo.

cd webapp
start "" http://127.0.0.1:5000/
python app.py

echo.
echo Server stopped.
pause
