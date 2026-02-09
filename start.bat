@echo off
echo Starting Statistical Calculator...
echo.
cd /d "%~dp0"

:: Check if .venv exists
if not exist ".venv" (
    echo First-time setup required. Running setup...
    call Setup_Windows.bat
    exit /b
)

:: Activate and run
call .venv\Scripts\activate.bat
streamlit run "import streamlit as st.py" --server.port 5180 --browser.gatherUsageStats false
