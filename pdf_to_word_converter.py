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
import pdfplumber
import fitz  # PyMuPDF
import io
import re
from typing import Optional, List, Tuple
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
        Extract text content from PDF file using multiple methods for best results

        Args:
            pdf_path (str): Path to the PDF file

        Returns:
            str: Extracted text content with proper spacing
        """
        extraction_methods = [
            ("pdfplumber", self._extract_with_pdfplumber),
            ("PyMuPDF", self._extract_with_pymupdf),
            ("PyPDF2", self._extract_with_pypdf2)
        ]

        for method_name, extraction_func in extraction_methods:
            try:
                logger.info(f"Trying extraction with {method_name}...")
                text_content = extraction_func(pdf_path)

                if text_content and text_content.strip():
                    logger.info(f"Successfully extracted text using {method_name}")

                    # Apply preprocessing to fix spacing issues
                    processed_text = self._preprocess_spacing(text_content)
                    return processed_text
                else:
                    logger.warning(f"{method_name} returned empty text")

            except Exception as e:
                logger.warning(f"{method_name} extraction failed: {str(e)}")
                continue

        raise Exception("All PDF extraction methods failed")

    def _extract_with_pdfplumber(self, pdf_path: str) -> str:
        """Extract text using pdfplumber (best for spacing)"""
        text_content = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text_content += page_text + "\n\n"
            logger.info(f"pdfplumber: Extracted text from {len(pdf.pages)} pages")
        return text_content

    def _extract_with_pymupdf(self, pdf_path: str) -> str:
        """Extract text using PyMuPDF (good for complex layouts)"""
        text_content = ""
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_text = page.get_text()
            if page_text:
                text_content += page_text + "\n\n"
        doc.close()
        logger.info(f"PyMuPDF: Extracted text from {len(doc)} pages")
        return text_content

    def _extract_with_pypdf2(self, pdf_path: str) -> str:
        """Extract text using PyPDF2 (fallback method)"""
        text_content = ""
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    text_content += page_text + "\n\n"
            logger.info(f"PyPDF2: Extracted text from {len(pdf_reader.pages)} pages")
        return text_content

    def _preprocess_spacing(self, text: str) -> str:
        """
        Intelligent preprocessing to fix common spacing issues before AI processing
        """
        logger.info("Applying intelligent spacing preprocessing...")

        # Fix merged words with common patterns
        processed_text = text

        # Pattern 1: Add space before capital letters in merged words (camelCase)
        # But avoid breaking acronyms or proper nouns
        processed_text = re.sub(r'([a-z])([A-Z])', r'\1 \2', processed_text)

        # Pattern 2: Add space between word and number
        processed_text = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', processed_text)
        processed_text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', processed_text)

        # Pattern 3: Fix common merged words (add more as needed)
        common_merges = {
            r'andthe': 'and the',
            r'ofthe': 'of the',
            r'inthe': 'in the',
            r'tothe': 'to the',
            r'forthe': 'for the',
            r'withthe': 'with the',
            r'onthe': 'on the',
            r'atthe': 'at the',
            r'bythe': 'by the',
            r'fromthe': 'from the',
            r'thatthe': 'that the',
            r'thisthe': 'this the',
            r'itis': 'it is',
            r'thisis': 'this is',
            r'thereis': 'there is',
            r'therefor': 'therefore',
            r'however': 'however',
            r'moreover': 'moreover',
            r'furthermore': 'furthermore',
        }

        for merged, separated in common_merges.items():
            processed_text = re.sub(merged, separated, processed_text, flags=re.IGNORECASE)

        # Pattern 4: Fix punctuation spacing
        processed_text = re.sub(r'([.!?])([A-Z])', r'\1 \2', processed_text)
        processed_text = re.sub(r'([,;:])([a-zA-Z])', r'\1 \2', processed_text)

        # Pattern 5: Clean up multiple spaces
        processed_text = re.sub(r' +', ' ', processed_text)

        # Pattern 6: Fix line breaks that split words
        processed_text = re.sub(r'([a-z])-\n([a-z])', r'\1\2', processed_text)

        logger.info("Spacing preprocessing completed")
        return processed_text

    def process_with_gemini(self, text_content: str) -> str:
        """
        Process extracted text with Gemini API for better formatting with model fallback

        Args:
            text_content (str): Raw text extracted from PDF

        Returns:
            str: Processed and formatted text
        """
        prompt = f"""
You are an expert text processor. Clean up this PDF-extracted text with EXTREME attention to word spacing:

CRITICAL SPACING RULES:
1. **MERGED WORDS**: Look for words stuck together and separate them properly
   - Example: "thequick" → "the quick"
   - Example: "andthe" → "and the"
   - Example: "itis" → "it is"
   - Example: "therefor" → "therefore"

2. **MISSING SPACES**: Add spaces where words are clearly merged
   - Between articles and nouns: "thedog" → "the dog"
   - Between prepositions: "inthe" → "in the"
   - Before capital letters: "wordAnother" → "word Another"
   - Between numbers and words: "5years" → "5 years"

3. **PRESERVE CORRECT SPACING**: Don't add extra spaces where they already exist correctly

4. **OTHER FIXES**:
   - Fix obvious OCR errors
   - Maintain paragraph structure
   - Fix punctuation spacing: "word.Another" → "word. Another"
   - Remove random line breaks within sentences
   - Keep original meaning and structure

5. **DO NOT**:
   - Add titles or headers
   - Change the document structure
   - Add commentary or explanations

EXAMPLES OF SPACING FIXES:
- "Thisisanimportantdocument" → "This is an important document"
- "Thecompanyreported5milliondollars" → "The company reported 5 million dollars"
- "However,theresults" → "However, the results"

Text to process:
{text_content}

Return ONLY the corrected text with proper spacing:

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

    def _post_process_spacing(self, text: str) -> str:
        """
        Final validation and fixing of spacing issues after AI processing
        """
        logger.info("Applying post-processing spacing validation...")

        processed_text = text

        # Check for remaining merged words using common patterns
        suspicious_patterns = [
            r'[a-z][A-Z]',  # camelCase
            r'[a-zA-Z]\d',  # word followed by number
            r'\d[a-zA-Z]',  # number followed by word
            r'[.!?][A-Z]',  # punctuation followed by capital
            r'[,;:][a-zA-Z]',  # punctuation followed by letter
        ]

        issues_found = 0
        for pattern in suspicious_patterns:
            matches = re.findall(pattern, processed_text)
            if matches:
                issues_found += len(matches)

        if issues_found > 0:
            logger.warning(f"Found {issues_found} potential spacing issues in post-processing")

            # Apply final fixes
            processed_text = re.sub(r'([a-z])([A-Z])', r'\1 \2', processed_text)
            processed_text = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', processed_text)
            processed_text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', processed_text)
            processed_text = re.sub(r'([.!?])([A-Z])', r'\1 \2', processed_text)
            processed_text = re.sub(r'([,;:])([a-zA-Z])', r'\1 \2', processed_text)

            # Clean up multiple spaces
            processed_text = re.sub(r' +', ' ', processed_text)

            logger.info("Applied final spacing corrections")
        else:
            logger.info("No spacing issues detected in post-processing")

        return processed_text

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
        ai_processed_text = self.process_with_gemini(raw_text)

        # Step 3: Post-process for final spacing validation
        logger.info("Applying final spacing validation...")
        final_text = self._post_process_spacing(ai_processed_text)

        # Step 4: Create Word document
        logger.info("Creating Word document...")
        self.create_word_document(final_text, output_path)
        
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
