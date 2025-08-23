@echo off
echo ========================================
echo  Entered On Audit System Launcher
echo ========================================
echo.

REM Change to the directory where the script is located
cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

echo Python found: 
python --version

REM Check if streamlit is installed
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Streamlit not found. Installing dependencies...
    echo Installing from requirements.txt...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
    echo Dependencies installed successfully!
) else (
    echo Streamlit found!
)

echo.
echo ========================================
echo  Starting Entered On Audit System...
echo ========================================
echo.
echo The app will open in your default browser
echo Press Ctrl+C in this window to stop the app
echo.

REM Launch the Streamlit app
streamlit run streamlit_app.py

REM If we reach here, the app has closed
echo.
echo ========================================
echo  App has closed
echo ========================================
pause