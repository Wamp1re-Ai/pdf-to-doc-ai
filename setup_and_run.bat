@echo off
echo ========================================
echo PDF to Word Converter Setup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

echo ✅ Python is installed
echo.

REM Install dependencies
echo Installing required packages...
pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ Failed to install dependencies
    pause
    exit /b 1
)

echo ✅ Dependencies installed successfully
echo.

REM Check for API key
if "%GEMINI_API_KEY%"=="" (
    echo ⚠️  GEMINI_API_KEY environment variable is not set
    echo.
    set /p api_key="Please enter your Gemini API key: "
    set GEMINI_API_KEY=%api_key%
    echo.
    echo ✅ API key set for this session
    echo To set it permanently, run:
    echo    setx GEMINI_API_KEY "%api_key%"
    echo.
) else (
    echo ✅ GEMINI_API_KEY is already set
)

echo.
echo ========================================
echo Ready to convert PDF files!
echo ========================================
echo.
echo Usage examples:
echo   python pdf_to_word_converter.py input.pdf
echo   python pdf_to_word_converter.py input.pdf -o output.docx
echo.

REM Ask if user wants to convert a file now
set /p convert="Do you want to convert a PDF file now? (y/n): "
if /i "%convert%"=="y" (
    set /p pdf_file="Enter path to PDF file: "
    if exist "%pdf_file%" (
        echo Converting "%pdf_file%"...
        python pdf_to_word_converter.py "%pdf_file%"
    ) else (
        echo ❌ File "%pdf_file%" not found
    )
)

echo.
pause
