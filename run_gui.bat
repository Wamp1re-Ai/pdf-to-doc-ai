@echo off
echo Starting PDF to Word Converter GUI...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

REM Check if required files exist
if not exist "pdf_to_word_converter.py" (
    echo ❌ pdf_to_word_converter.py not found
    echo Please ensure all files are in the same directory
    pause
    exit /b 1
)

if not exist "pdf_to_word_gui.py" (
    echo ❌ pdf_to_word_gui.py not found
    echo Please ensure all files are in the same directory
    pause
    exit /b 1
)

REM Install dependencies if needed
echo Checking dependencies...
pip show google-generativeai >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ Failed to install dependencies
        pause
        exit /b 1
    )
)

echo ✅ All dependencies are installed
echo.

REM Launch the GUI
echo 🚀 Launching PDF to Word Converter GUI...
python pdf_to_word_gui.py

pause
