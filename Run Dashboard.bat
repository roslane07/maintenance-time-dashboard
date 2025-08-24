@echo off
SETLOCAL

:: Enable command output
echo Starting dashboard setup...

:: Check if Python is installed
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not in PATH
    pause
    exit /b 1
)

:: Create virtual environment if it doesn't exist
if not exist ".venv\Scripts\activate.bat" (
    echo Creating virtual environment...
    python -m venv .venv
)

:: Activate virtual environment
call .venv\Scripts\activate.bat

:: Install or upgrade pip
pip install --upgrade pip

:: Install requirements
pip install -r app\requirements.txt

:: Launch the dashboard
cd app
streamlit run main.py

:: Keep window open if there's an error
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to start dashboard
    pause
)
ENDLOCAL
