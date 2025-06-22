@echo off
echo ========================================
echo PDF to Word Converter - Complete Setup
echo ========================================
echo.

REM Check if we're in the right directory
echo Current directory: %CD%
echo.

REM Check if this is a fresh download
if not exist "pdf_to_word_converter.py" (
    echo ❌ Setup files not found!
    echo.
    echo This appears to be a fresh setup. Please:
    echo.
    echo 1. Go to: https://github.com/Wamp1re-Ai/pdf-to-doc-ai
    echo 2. Click the green "Code" button
    echo 3. Select "Download ZIP"
    echo 4. Extract ALL files to a folder
    echo 5. Run this setup script from that folder
    echo.
    echo OR clone with git:
    echo git clone https://github.com/Wamp1re-Ai/pdf-to-doc-ai.git
    echo.
    pause
    exit /b 1
)

echo ✅ Found main converter file
echo.

REM List all files to verify complete download
echo Files found in current directory:
echo ----------------------------------------
dir *.py *.txt *.bat *.md /b 2>nul
echo ----------------------------------------
echo.

REM Check for all required files
set missing_files=0

if not exist "pdf_to_word_converter.py" (
    echo ❌ pdf_to_word_converter.py missing
    set /a missing_files+=1
)

if not exist "pdf_to_word_gui.py" (
    echo ❌ pdf_to_word_gui.py missing
    set /a missing_files+=1
)

if not exist "requirements.txt" (
    echo ❌ requirements.txt missing
    set /a missing_files+=1
)

if not exist "README.md" (
    echo ⚠️  README.md missing (optional)
)

if %missing_files% gtr 0 (
    echo.
    echo ❌ %missing_files% required files are missing!
    echo Please download the complete repository from:
    echo https://github.com/Wamp1re-Ai/pdf-to-doc-ai
    echo.
    pause
    exit /b 1
)

echo ✅ All required files found!
echo.

REM Check Python
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python not found!
    echo.
    echo Please install Python 3.7+ from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
) else (
    echo ✅ Python found:
    python --version
)
echo.

REM Install all dependencies
echo Installing all required packages...
echo ========================================

pip install google-generativeai python-docx PyPDF2 pdfplumber PyMuPDF

if errorlevel 1 (
    echo.
    echo ❌ Some packages failed to install
    echo Trying with requirements.txt...
    pip install -r requirements.txt
    
    if errorlevel 1 (
        echo.
        echo ❌ Installation failed!
        echo.
        echo Please try manually:
        echo pip install google-generativeai
        echo pip install python-docx  
        echo pip install PyPDF2
        echo pip install pdfplumber
        echo pip install PyMuPDF
        echo.
        pause
        exit /b 1
    )
)

echo.
echo ✅ All packages installed successfully!
echo.

REM Test the installation
echo Testing installation...
echo ========================================

python -c "import pdf_to_word_converter; print('✅ Main converter module works')" 2>nul
if errorlevel 1 (
    echo ❌ Main converter has issues
    python -c "import pdf_to_word_converter"
    pause
    exit /b 1
)

python -c "import pdf_to_word_gui; print('✅ GUI module works')" 2>nul
if errorlevel 1 (
    echo ❌ GUI module has issues  
    python -c "import pdf_to_word_gui"
    pause
    exit /b 1
)

echo.
echo ========================================
echo 🎉 SETUP COMPLETE! 🎉
echo ========================================
echo.
echo Your PDF to Word Converter is ready to use!
echo.
echo Next steps:
echo 1. Get a Gemini API key from: https://makersuite.google.com/app/apikey
echo 2. Double-click 'run_gui.bat' to start the converter
echo 3. Enter your API key and convert PDFs!
echo.
echo Files you can use:
echo - run_gui.bat          (Start the GUI)
echo - pdf_to_word_gui.py   (Direct GUI launch)
echo - pdf_to_word_converter.py (Command line)
echo.
pause
