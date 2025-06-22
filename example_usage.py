#!/usr/bin/env python3
"""
Example usage of the PDF to Word converter
"""

import os
from pdf_to_word_converter import PDFToWordConverter

def example_conversion():
    """Example of how to use the PDFToWordConverter class"""
    
    # Get API key from environment variable
    api_key = os.getenv('GEMINI_API_KEY')
    
    if not api_key:
        print("❌ Please set GEMINI_API_KEY environment variable")
        print("   Windows: set GEMINI_API_KEY=your_api_key_here")
        print("   Linux/Mac: export GEMINI_API_KEY=your_api_key_here")
        return
    
    # Initialize converter
    converter = PDFToWordConverter(api_key)
    
    # Example PDF file (replace with your actual PDF file)
    pdf_file = "sample.pdf"
    
    if not os.path.exists(pdf_file):
        print(f"❌ PDF file '{pdf_file}' not found")
        print("   Please place a PDF file named 'sample.pdf' in this directory")
        return
    
    try:
        # Convert PDF to Word
        output_file = converter.convert_pdf_to_word(
            pdf_path=pdf_file,
            output_path="converted_document.docx"
        )
        
        print(f"✅ Conversion successful!")
        print(f"📄 Input: {pdf_file}")
        print(f"📝 Output: {output_file}")
        
    except Exception as e:
        print(f"❌ Conversion failed: {str(e)}")

if __name__ == "__main__":
    example_conversion()
