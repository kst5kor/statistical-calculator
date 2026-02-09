@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Statistical Process Capability Tool
echo   Windows Setup Script
echo ============================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH
    echo.
    echo Please install Python 3.10+ from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

:: Check if virtual environment exists
if not exist ".venv" (
    echo [INFO] Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
) else (
    echo [OK] Virtual environment already exists
)

:: Activate virtual environment
echo [INFO] Activating virtual environment...
call .venv\Scripts\activate.bat

:: Install/upgrade pip
echo [INFO] Upgrading pip...
python -m pip install --upgrade pip --quiet

:: Install requirements
echo [INFO] Installing dependencies (this may take a few minutes)...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed

echo.
echo ============================================
echo   Setup Complete!
echo ============================================
echo.
echo Starting the application...
echo Your browser will open automatically.
echo.
echo To close the app, press Ctrl+C in this window
echo or simply close this window.
echo.
echo ============================================
echo.

:: Start streamlit
streamlit run "import streamlit as st.py" --server.port 5180 --browser.gatherUsageStats false

pause
