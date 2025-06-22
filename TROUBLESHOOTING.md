# Troubleshooting Guide

## Common Issues and Solutions

### 1. "pdf_to_word_converter.py not found"

**Problem:** The batch file can't find the required Python files.

**Solutions:**
1. **Download Complete Repository:**
   - Go to: https://github.com/Wamp1re-Ai/pdf-to-doc-ai
   - Click green "Code" button → "Download ZIP"
   - Extract ALL files to the same folder

2. **Check File Location:**
   - Make sure all files are in the same directory
   - Required files: `pdf_to_word_converter.py`, `pdf_to_word_gui.py`, `requirements.txt`

3. **Run from Correct Directory:**
   - Open Command Prompt in the folder containing the Python files
   - Or double-click the batch file from the correct folder

### 2. "Python is not installed or not in PATH"

**Problem:** Python is not installed or not accessible.

**Solutions:**
1. **Install Python:**
   - Download from: https://www.python.org/downloads/
   - **IMPORTANT:** Check "Add Python to PATH" during installation

2. **Verify Installation:**
   ```cmd
   python --version
   ```

3. **Manual PATH Setup (if needed):**
   - Add Python installation directory to Windows PATH
   - Usually: `C:\Users\[username]\AppData\Local\Programs\Python\Python3x\`

### 3. "Failed to install dependencies"

**Problem:** Package installation fails.

**Solutions:**
1. **Update pip:**
   ```cmd
   python -m pip install --upgrade pip
   ```

2. **Install packages individually:**
   ```cmd
   pip install google-generativeai
   pip install python-docx
   pip install PyPDF2
   pip install pdfplumber
   pip install PyMuPDF
   ```

3. **Use user installation:**
   ```cmd
   pip install --user -r requirements.txt
   ```

### 4. "Module not found" errors

**Problem:** Python can't find installed packages.

**Solutions:**
1. **Check installation:**
   ```cmd
   pip list | findstr google-generativeai
   ```

2. **Reinstall specific package:**
   ```cmd
   pip uninstall [package-name]
   pip install [package-name]
   ```

3. **Use virtual environment:**
   ```cmd
   python -m venv pdf_converter
   pdf_converter\Scripts\activate
   pip install -r requirements.txt
   ```

### 5. "Gemini API key is required"

**Problem:** No API key provided.

**Solutions:**
1. **Get API Key:**
   - Go to: https://makersuite.google.com/app/apikey
   - Sign in with Google account
   - Create new API key

2. **Set Environment Variable (Recommended):**
   ```cmd
   setx GEMINI_API_KEY "your_api_key_here"
   ```

3. **Enter in GUI:**
   - Paste API key in the GUI field
   - Click the "Help" button for instructions

### 6. GUI doesn't start

**Problem:** GUI window doesn't appear.

**Solutions:**
1. **Run from Command Prompt:**
   ```cmd
   python pdf_to_word_gui.py
   ```
   - Check for error messages

2. **Check tkinter:**
   ```cmd
   python -c "import tkinter; print('tkinter OK')"
   ```

3. **Try alternative launch:**
   ```cmd
   python -m pdf_to_word_gui
   ```

### 7. "All PDF extraction methods failed"

**Problem:** Can't extract text from PDF.

**Solutions:**
1. **Check PDF file:**
   - Ensure PDF is not corrupted
   - Try with a different PDF file

2. **PDF might be image-based:**
   - Use OCR software first to convert to text-based PDF
   - Try Adobe Acrobat or online OCR tools

3. **Check file permissions:**
   - Ensure PDF is not password-protected
   - Check file is not in use by another program

### 8. "Conversion failed" with Gemini API

**Problem:** AI processing fails.

**Solutions:**
1. **Check API key:**
   - Verify key is correct and active
   - Check quota limits in Google AI Studio

2. **Check internet connection:**
   - Ensure stable internet connection
   - Try again after a few minutes

3. **Try different model:**
   - The converter automatically tries fallback models
   - Check logs for which models were attempted

## Quick Setup Commands

### Complete Fresh Setup:
```cmd
# 1. Download repository
git clone https://github.com/Wamp1re-Ai/pdf-to-doc-ai.git
cd pdf-to-doc-ai

# 2. Install dependencies
pip install -r requirements.txt

# 3. Set API key
setx GEMINI_API_KEY "your_api_key_here"

# 4. Run GUI
python pdf_to_word_gui.py
```

### Manual Dependency Installation:
```cmd
pip install google-generativeai python-docx PyPDF2 pdfplumber PyMuPDF
```

### Test Installation:
```cmd
python -c "import pdf_to_word_converter; print('✅ Ready to use!')"
```

## Getting Help

If you're still having issues:

1. **Check the error message carefully**
2. **Try the setup_complete.bat script** for automated setup
3. **Run from Command Prompt** to see detailed error messages
4. **Check the GitHub repository** for updates: https://github.com/Wamp1re-Ai/pdf-to-doc-ai

## System Requirements

- **Python 3.7+** (3.8+ recommended)
- **Windows 10+** (for batch files)
- **Internet connection** (for Gemini API)
- **Gemini API key** (free from Google AI Studio)

## File Structure

Your folder should contain:
```
pdf-to-doc-ai/
├── pdf_to_word_converter.py    # Main converter
├── pdf_to_word_gui.py          # GUI application  
├── requirements.txt            # Dependencies
├── run_gui.bat                 # GUI launcher
├── setup_complete.bat          # Complete setup
├── README.md                   # Documentation
└── TROUBLESHOOTING.md          # This file
```
