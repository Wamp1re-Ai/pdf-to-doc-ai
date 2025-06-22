#!/usr/bin/env python3
"""
PDF to Word Converter using Gemini API
This script converts PDF files to Word documents using Google's Gemini API
for intelligent text processing and formatting.
"""

import os
import sys
import argparse
from pathlib import Path
import google.generativeai as genai
from docx import Document
from docx.shared import Inches
import PyPDF2
import io
from typing import Optional, List
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFToWordConverter:
    def __init__(self, api_key: str):
        """
        Initialize the converter with Gemini API key

        Args:
            api_key (str): Google Gemini API key
        """
        self.api_key = api_key
        genai.configure(api_key=api_key)

        # Model fallback order: Flash 2.5 -> Flash 2.0 -> Gemma 27B
        self.model_names = [
            'gemini-2.0-flash-exp',  # Latest Flash 2.5
            'gemini-1.5-flash',      # Flash 2.0
            'gemini-1.5-pro',        # Pro model as backup
            'gemini-pro'             # Original pro model
        ]
        self.current_model = None
        self._initialize_model()

    def _initialize_model(self):
        """Initialize the best available model with fallback"""
        for model_name in self.model_names:
            try:
                self.current_model = genai.GenerativeModel(model_name)
                logger.info(f"Successfully initialized model: {model_name}")
                break
            except Exception as e:
                logger.warning(f"Failed to initialize {model_name}: {str(e)}")
                continue

        if self.current_model is None:
            raise Exception("Failed to initialize any Gemini model")

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """
        Extract text content from PDF file
        
        Args:
            pdf_path (str): Path to the PDF file
            
        Returns:
            str: Extracted text content
        """
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text_content = ""
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    text_content += page.extract_text() + "\n\n"
                    
                logger.info(f"Extracted text from {len(pdf_reader.pages)} pages")
                return text_content
                
        except Exception as e:
            logger.error(f"Error extracting text from PDF: {str(e)}")
            raise
    
    def process_with_gemini(self, text_content: str) -> str:
        """
        Process extracted text with Gemini API for better formatting with model fallback

        Args:
            text_content (str): Raw text extracted from PDF

        Returns:
            str: Processed and formatted text
        """
        prompt = f"""
Clean up and format this text extracted from a PDF document. Follow these rules strictly:

1. Fix OCR errors and garbled text
2. Preserve original spacing and paragraph structure
3. Maintain exact character spacing - do not add extra spaces between words
4. Keep the original meaning and content structure
5. Remove only obvious formatting artifacts (like random line breaks mid-sentence)
6. Ensure proper punctuation and grammar
7. Do NOT add any titles, headers, or introductory text
8. Return ONLY the cleaned original content

Text to process:
{text_content}

Return only the cleaned text with no additional commentary or formatting."""

        # Try models in fallback order
        for i, model_name in enumerate(self.model_names):
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                logger.info(f"Text processed successfully with {model_name}")
                return response.text

            except Exception as e:
                logger.warning(f"Failed to process with {model_name}: {str(e)}")
                if i < len(self.model_names) - 1:
                    logger.info(f"Falling back to next model...")
                    continue
                else:
                    logger.error("All models failed, returning original text")
                    return text_content

        return text_content
    
    def create_word_document(self, processed_text: str, output_path: str) -> None:
        """
        Create a Word document from processed text with better formatting

        Args:
            processed_text (str): Formatted text content
            output_path (str): Path for the output Word document
        """
        try:
            doc = Document()

            # Do NOT add a title - start directly with content

            # Split text into paragraphs and add to document
            paragraphs = processed_text.split('\n\n')

            for paragraph_text in paragraphs:
                if paragraph_text.strip():
                    # Preserve original spacing and formatting
                    lines = paragraph_text.split('\n')

                    for line in lines:
                        if line.strip():
                            # Check if it looks like a heading (short line, all caps, or starts with numbers/bullets)
                            line_stripped = line.strip()
                            is_heading = (
                                len(line_stripped) < 80 and
                                (line_stripped.isupper() or
                                 line_stripped.startswith(('1.', '2.', '3.', '4.', '5.', 'Chapter', 'CHAPTER', 'Section', 'SECTION')) or
                                 line_stripped.endswith(':'))
                            )

                            if is_heading:
                                doc.add_heading(line_stripped, level=1)
                            else:
                                # Add as regular paragraph, preserving spacing
                                paragraph = doc.add_paragraph()
                                paragraph.add_run(line.rstrip())  # Remove only trailing whitespace
                        else:
                            # Add empty line for spacing
                            doc.add_paragraph()

            # Save the document
            doc.save(output_path)
            logger.info(f"Word document saved to: {output_path}")

        except Exception as e:
            logger.error(f"Error creating Word document: {str(e)}")
            raise
    
    def convert_pdf_to_word(self, pdf_path: str, output_path: Optional[str] = None) -> str:
        """
        Main conversion method
        
        Args:
            pdf_path (str): Path to input PDF file
            output_path (str, optional): Path for output Word file
            
        Returns:
            str: Path to the created Word document
        """
        # Validate input file
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        # Generate output path if not provided
        if output_path is None:
            pdf_name = Path(pdf_path).stem
            output_path = f"{pdf_name}_converted.docx"
        
        logger.info(f"Starting conversion: {pdf_path} -> {output_path}")
        
        # Step 1: Extract text from PDF
        logger.info("Extracting text from PDF...")
        raw_text = self.extract_text_from_pdf(pdf_path)
        
        if not raw_text.strip():
            raise ValueError("No text content found in PDF")
        
        # Step 2: Process with Gemini API
        logger.info("Processing text with Gemini API...")
        processed_text = self.process_with_gemini(raw_text)
        
        # Step 3: Create Word document
        logger.info("Creating Word document...")
        self.create_word_document(processed_text, output_path)
        
        logger.info("Conversion completed successfully!")
        return output_path

def main():
    """Main function to handle command line arguments and run conversion"""
    parser = argparse.ArgumentParser(description='Convert PDF to Word using Gemini API')
    parser.add_argument('pdf_file', help='Path to the PDF file to convert')
    parser.add_argument('-o', '--output', help='Output Word file path (optional)')
    parser.add_argument('-k', '--api-key', help='Gemini API key (or set GEMINI_API_KEY env var)')
    
    args = parser.parse_args()
    
    # Get API key from argument or environment variable
    api_key = args.api_key or os.getenv('GEMINI_API_KEY')
    
    if not api_key:
        logger.error("Gemini API key is required. Provide it via --api-key argument or GEMINI_API_KEY environment variable")
        sys.exit(1)
    
    try:
        # Initialize converter
        converter = PDFToWordConverter(api_key)
        
        # Perform conversion
        output_file = converter.convert_pdf_to_word(args.pdf_file, args.output)
        
        print(f"✅ Conversion successful!")
        print(f"📄 Input: {args.pdf_file}")
        print(f"📝 Output: {output_file}")
        
    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
