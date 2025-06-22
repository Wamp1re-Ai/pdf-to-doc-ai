#!/usr/bin/env python3
"""
Test script for the PDF to Word converter
Tests the model fallback and formatting improvements
"""

import os
import sys
from pdf_to_word_converter import PDFToWordConverter

def test_model_initialization():
    """Test model initialization and fallback"""
    print("Testing model initialization...")
    
    api_key = os.getenv('GEMINI_API_KEY')
    if not api_key:
        print("❌ GEMINI_API_KEY not set")
        return False
    
    try:
        converter = PDFToWordConverter(api_key)
        print(f"✅ Successfully initialized converter")
        print(f"📋 Available models: {converter.model_names}")
        return True
    except Exception as e:
        print(f"❌ Failed to initialize converter: {e}")
        return False

def test_text_processing():
    """Test text processing with sample text"""
    print("\nTesting text processing...")
    
    api_key = os.getenv('GEMINI_API_KEY')
    if not api_key:
        print("❌ GEMINI_API_KEY not set")
        return False
    
    # Sample text with spacing issues
    sample_text = """This   is   a   test   document   with   irregular   spacing.

CHAPTER 1: INTRODUCTION

This paragraph has  multiple  spaces  between  words  that  should  be  preserved  properly.

Some bullet points:
1. First point with proper spacing
2. Second point  with  extra  spaces
3. Third point

CONCLUSION:
The document should maintain proper formatting without adding unnecessary titles."""

    try:
        converter = PDFToWordConverter(api_key)
        processed_text = converter.process_with_gemini(sample_text)
        
        print("✅ Text processing successful")
        print("\n📄 Original text preview:")
        print(sample_text[:200] + "...")
        print("\n📝 Processed text preview:")
        print(processed_text[:200] + "...")
        
        # Check if unnecessary title was added
        if processed_text.lower().startswith('converted document'):
            print("⚠️  Warning: Unnecessary title detected")
        else:
            print("✅ No unnecessary title added")
            
        return True
        
    except Exception as e:
        print(f"❌ Text processing failed: {e}")
        return False

def test_document_creation():
    """Test Word document creation"""
    print("\nTesting document creation...")
    
    sample_text = """SAMPLE DOCUMENT

This is a test paragraph with proper spacing and formatting.

Another paragraph with different content to test the document structure."""

    try:
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            print("❌ GEMINI_API_KEY not set")
            return False
            
        converter = PDFToWordConverter(api_key)
        test_output = "test_output.docx"
        
        converter.create_word_document(sample_text, test_output)
        
        if os.path.exists(test_output):
            print(f"✅ Document created successfully: {test_output}")
            print("📄 You can open the document to verify formatting")
            return True
        else:
            print("❌ Document file not found")
            return False
            
    except Exception as e:
        print(f"❌ Document creation failed: {e}")
        return False

def main():
    """Run all tests"""
    print("🧪 PDF to Word Converter Test Suite")
    print("=" * 50)
    
    tests = [
        test_model_initialization,
        test_text_processing,
        test_document_creation
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
        print("-" * 30)
    
    print(f"\n📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed!")
    else:
        print("⚠️  Some tests failed. Check the output above.")

if __name__ == "__main__":
    main()
