# PDF to Word Converter using Gemini API

This tool converts PDF files to Word documents using Google's Gemini API for intelligent text processing and formatting. Available as both a command-line script and a user-friendly GUI application.

## Features

- 📄 **Multi-Method PDF Extraction** (pdfplumber → PyMuPDF → PyPDF2)
- 🔧 **Intelligent Spacing Detection** - Fixes merged words automatically
- 🤖 **Advanced AI Processing** with model fallback (Flash 2.5 → Flash 2.0 → Pro)
- 📝 Generate clean Word documents without unnecessary titles
- 🖥️ **Simple GUI application** for easy drag-and-drop usage
- 🔧 Command-line interface for automation and scripting
- 📋 **4-Layer Processing**: Extraction → Preprocessing → AI → Post-processing
- 🚀 One-click setup and conversion
- 🔄 **Automatic fallback** for maximum reliability
- ✨ **Perfect spacing preservation** - No more merged words!

## Prerequisites

1. **Python 3.7+** installed on your system
2. **Google Gemini API key** - Get one from [Google AI Studio](https://makersuite.google.com/app/apikey)

## Installation

1. Clone or download this repository
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Setup

### Option 1: Environment Variable (Recommended)
Set your Gemini API key as an environment variable:

**Windows:**
```cmd
set GEMINI_API_KEY=your_api_key_here
```

**Linux/Mac:**
```bash
export GEMINI_API_KEY=your_api_key_here
```

### Option 2: Command Line Argument
Pass the API key directly when running the script (see usage below).

## Usage

### GUI Application (Recommended for most users)

**Quick Start:**
1. Double-click `run_gui.bat` (Windows) or run `python pdf_to_word_gui.py`
2. Enter your Gemini API key
3. Click "Browse" to select your PDF file
4. Click "Convert PDF to Word"
5. Done! 🎉

**GUI Features:**
- 🖱️ Point-and-click interface
- 📁 File browser for easy file selection
- 📊 Real-time progress tracking
- 📝 Conversion log with detailed status
- ❓ Built-in help for API key setup
- 🔄 Automatic output file naming

### Command-Line Usage

### Basic Usage
```bash
python pdf_to_word_converter.py input.pdf
```

### Specify Output File
```bash
python pdf_to_word_converter.py input.pdf -o output.docx
```

### Provide API Key via Command Line
```bash
python pdf_to_word_converter.py input.pdf -k your_api_key_here
```

### Full Example
```bash
python pdf_to_word_converter.py document.pdf -o converted_document.docx -k your_api_key_here
```

## Command Line Arguments

- `pdf_file` (required): Path to the PDF file to convert
- `-o, --output` (optional): Output Word file path. If not specified, uses `{input_name}_converted.docx`
- `-k, --api-key` (optional): Gemini API key. Can also be set via `GEMINI_API_KEY` environment variable

## How It Works - 4-Layer Processing System

### 🔄 **Layer 1: Multi-Method PDF Extraction**
- **Primary**: `pdfplumber` (best spacing preservation)
- **Fallback 1**: `PyMuPDF` (excellent for complex layouts)
- **Fallback 2**: `PyPDF2` (reliable baseline)
- Automatically selects the best extraction method

### 🧠 **Layer 2: Intelligent Preprocessing**
- **Merged Word Detection**: Identifies words stuck together
- **Pattern Recognition**: Fixes common issues like "andthe" → "and the"
- **Smart Spacing**: Adds spaces before capitals, after punctuation
- **Number Separation**: "5years" → "5 years"

### 🤖 **Layer 3: AI Enhancement**
Uses advanced Gemini models with automatic fallback:
- **Primary**: Gemini 2.0 Flash (latest and fastest)
- **Fallback 1**: Gemini 1.5 Flash (reliable alternative)
- **Fallback 2**: Gemini 1.5 Pro (comprehensive processing)
- **Final**: Gemini Pro (stable baseline)

**AI Tasks**:
- OCR error correction with spacing focus
- Grammar and punctuation fixes
- Document structure preservation
- NO artificial titles or headers added

### ✅ **Layer 4: Post-Processing Validation**
- **Final Spacing Check**: Scans for remaining merged words
- **Pattern Validation**: Ensures proper word boundaries
- **Quality Assurance**: Applies final corrections if needed

### 📝 **Layer 5: Document Creation**
- Uses python-docx to create clean, properly formatted documents
- Preserves original structure and spacing
- Smart heading detection without artificial additions

## Example Output

The script will provide progress updates:
```
2024-01-15 10:30:15 - INFO - Extracting text from PDF...
2024-01-15 10:30:16 - INFO - Extracted text from 5 pages
2024-01-15 10:30:16 - INFO - Processing text with Gemini API...
2024-01-15 10:30:18 - INFO - Text processed successfully with Gemini API
2024-01-15 10:30:18 - INFO - Creating Word document...
2024-01-15 10:30:18 - INFO - Word document saved to: document_converted.docx
2024-01-15 10:30:18 - INFO - Conversion completed successfully!

✅ Conversion successful!
📄 Input: document.pdf
📝 Output: document_converted.docx
```

## Error Handling

The script includes robust error handling for:
- Missing or invalid PDF files
- API key issues
- Network connectivity problems
- Text extraction failures
- Word document creation errors

If Gemini API processing fails, the script will fall back to using the raw extracted text.

## Limitations

- PDF files with complex layouts or images may not convert perfectly
- Very large PDF files may take longer to process
- Gemini API has rate limits and usage costs
- Scanned PDFs (images) may have OCR limitations

## Troubleshooting

### Common Issues

1. **"No module named 'google.generativeai'"**
   - Run: `pip install -r requirements.txt`

2. **"Gemini API key is required"**
   - Set the `GEMINI_API_KEY` environment variable or use `-k` argument

3. **"PDF file not found"**
   - Check the file path and ensure the PDF exists

4. **"No text content found in PDF"**
   - The PDF might be image-based or corrupted

## License

This project is open source and available under the MIT License.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.
