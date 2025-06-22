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

def test_spacing_fixes():
    """Test spacing preprocessing and fixes"""
    print("\nTesting spacing fixes...")

    api_key = os.getenv('GEMINI_API_KEY')
    if not api_key:
        print("❌ GEMINI_API_KEY not set")
        return False

    # Sample text with common spacing issues
    test_cases = [
        ("thequickbrownfox", "the quick brown fox"),
        ("andtheresults", "and the results"),
        ("itis5years", "it is 5 years"),
        ("word.Another", "word. Another"),
        ("However,thecompany", "However, the company"),
        ("therefor", "therefore"),
        ("wordAnother", "word Another"),
    ]

    try:
        converter = PDFToWordConverter(api_key)

        print("🔧 Testing preprocessing fixes:")
        for original, expected in test_cases:
            fixed = converter._preprocess_spacing(original)
            if expected.lower() in fixed.lower():
                print(f"  ✅ '{original}' → '{fixed}'")
            else:
                print(f"  ⚠️  '{original}' → '{fixed}' (expected: '{expected}')")

        # Test full processing with merged words
        sample_text = """Thisisanimportantdocumentaboutthecompany.Theresultsshow5milliondollars.However,therearesomeissues."""

        processed_text = converter.process_with_gemini(sample_text)

        print("\n📄 Original merged text:")
        print(sample_text)
        print("\n📝 After AI processing:")
        print(processed_text)

        # Check for common spacing issues
        spacing_issues = [
            "thisis", "thecompany", "theresults", "thereare", "5million"
        ]

        issues_found = 0
        for issue in spacing_issues:
            if issue.lower() in processed_text.lower():
                issues_found += 1
                print(f"  ⚠️  Still found: '{issue}'")

        if issues_found == 0:
            print("✅ No spacing issues detected in processed text")
        else:
            print(f"⚠️  Found {issues_found} remaining spacing issues")

        return issues_found == 0

    except Exception as e:
        print(f"❌ Spacing test failed: {e}")
        return False

def test_text_processing():
    """Test text processing with sample text"""
    print("\nTesting comprehensive text processing...")

    api_key = os.getenv('GEMINI_API_KEY')
    if not api_key:
        print("❌ GEMINI_API_KEY not set")
        return False

    # Sample text with various issues
    sample_text = """Thisisatestdocumentwithspacingissues.

CHAPTER1:INTRODUCTION

Thisparagraphhasmultiplespacesbetweenwordsandmergedwords.Thecompanyreported5milliondollars.However,theresultswerenot satisfactory.

Somebulletpoints:
1.Firstpointwithproperspacingissues
2.Secondpointwithextraspaces
3.Thirdpoint

CONCLUSION:
Thedocumentshouldmaintainproperformattingwithoutaddingtitles."""

    try:
        converter = PDFToWordConverter(api_key)
        processed_text = converter.process_with_gemini(sample_text)

        print("✅ Text processing successful")
        print("\n📄 Original text preview:")
        print(sample_text[:150] + "...")
        print("\n📝 Processed text preview:")
        print(processed_text[:150] + "...")

        # Check if unnecessary title was added
        if processed_text.lower().startswith('converted document'):
            print("⚠️  Warning: Unnecessary title detected")
        else:
            print("✅ No unnecessary title added")

        # Check for spacing improvements
        if "This is a test" in processed_text and "The company reported" in processed_text:
            print("✅ Spacing improvements detected")
        else:
            print("⚠️  Spacing improvements may be incomplete")

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
        test_spacing_fixes,
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
