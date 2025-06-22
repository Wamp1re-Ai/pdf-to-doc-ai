@echo off
echo ========================================
echo PDF to Word Converter GUI Launcher
echo ========================================
echo.

REM Show current directory
echo Current directory: %CD%
echo.

REM List all Python files in current directory
echo Files in current directory:
dir *.py /b 2>nul
if errorlevel 1 (
    echo No Python files found in current directory!
    echo.
)
echo.

REM Check if Python is installed
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    echo.
    pause
    exit /b 1
) else (
    echo ✅ Python is installed
    python --version
)
echo.

REM Check if required files exist with detailed diagnostics
echo Checking required files...

if not exist "pdf_to_word_converter.py" (
    echo ❌ pdf_to_word_converter.py not found
    echo.
    echo Troubleshooting:
    echo 1. Make sure you downloaded ALL files from the repository
    echo 2. Extract all files to the same folder
    echo 3. Run this batch file from the same folder as the Python files
    echo.
    echo Expected files:
    echo - pdf_to_word_converter.py
    echo - pdf_to_word_gui.py
    echo - requirements.txt
    echo - run_gui.bat (this file)
    echo.
    pause
    exit /b 1
) else (
    echo ✅ pdf_to_word_converter.py found
)

if not exist "pdf_to_word_gui.py" (
    echo ❌ pdf_to_word_gui.py not found
    echo Please ensure all files are in the same directory
    echo.
    pause
    exit /b 1
) else (
    echo ✅ pdf_to_word_gui.py found
)

if not exist "requirements.txt" (
    echo ⚠️  requirements.txt not found
    echo Will try to install dependencies manually...
) else (
    echo ✅ requirements.txt found
)

echo.

REM Install dependencies
echo ========================================
echo Installing Dependencies
echo ========================================

if exist "requirements.txt" (
    echo Installing from requirements.txt...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ Failed to install from requirements.txt
        echo Trying manual installation...
        goto manual_install
    ) else (
        echo ✅ Dependencies installed from requirements.txt
        goto check_deps
    )
) else (
    echo requirements.txt not found, installing manually...
    goto manual_install
)

:manual_install
echo Installing dependencies manually...
pip install google-generativeai python-docx PyPDF2 pdfplumber PyMuPDF
if errorlevel 1 (
    echo ❌ Failed to install dependencies manually
    echo.
    echo Please try:
    echo 1. pip install google-generativeai
    echo 2. pip install python-docx
    echo 3. pip install PyPDF2
    echo 4. pip install pdfplumber
    echo 5. pip install PyMuPDF
    echo.
    pause
    exit /b 1
) else (
    echo ✅ Dependencies installed manually
)

:check_deps
echo.
echo Verifying installation...
python -c "import google.generativeai; print('✅ google-generativeai OK')" 2>nul || echo "❌ google-generativeai missing"
python -c "import docx; print('✅ python-docx OK')" 2>nul || echo "❌ python-docx missing"
python -c "import PyPDF2; print('✅ PyPDF2 OK')" 2>nul || echo "❌ PyPDF2 missing"
python -c "import pdfplumber; print('✅ pdfplumber OK')" 2>nul || echo "❌ pdfplumber missing"
python -c "import fitz; print('✅ PyMuPDF OK')" 2>nul || echo "❌ PyMuPDF missing"

echo.
echo ========================================
echo Launching GUI
echo ========================================
echo 🚀 Starting PDF to Word Converter GUI...
echo.

python pdf_to_word_gui.py

echo.
echo ========================================
echo GUI Closed
echo ========================================
pause
